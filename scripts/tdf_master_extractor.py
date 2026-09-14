"""実寸法師(.tdf)から製品マスタ情報(図番・製品マーク・設計符号・サイズ・本数・
重量・長さ・左継手・右継手)を抽出するロジック。

2026-09-11、YK-AB01(普通梁)・WR-1G-1(斜め梁)の2ファイルの実機解析で確立した
ルールをコード化したもの。適用範囲はこの2パターンの検証に基づく。

前提となる図面構造:
  - 製品情報は「製品マーク | 設計符号 | サイズ | 本数(N台) [| 重量(N.Nkg)]」という
    横並びテーブルとして、pad=0の通常テキストで描画されている。
  - 図番はタイトル欄のテキスト(高さH=100前後、他の注記より大きい)。
  - 「長さ」は、製品マーク〜本数のX範囲・製品マークのYより上にある数値寸法群の
    うち、(a)図面の主軸方向に沿う直線と長さが一致する本数が最多で、
    (b)出現回数が多すぎる値(取付ピッチ)は原則除外するが、他の系列との合計が
    一致するなら復活する、という手順で1つに絞り込む。
  - 「左継手・右継手」は、長丸(半径が一致し開始角の差が180度の円弧ペア)の
    中心に対応するテキストを候補とし、回転後座標で同じY(勾配)を持つペアを
    その製品の継手とみなし、X(回転後)の小さい方を左、大きい方を右とする。
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field

import tdf_binary as tb


# ---------------------------------------------------------------------------
# 共通ユーティリティ
# ---------------------------------------------------------------------------

def _try_float(s: str | None) -> float | None:
    if s is None:
        return None
    try:
        return float(s)
    except ValueError:
        return None


SIZE_PATTERN = re.compile(
    r"^(H|BH|SH|TH|\[|L|BOX|角|[0-9]+φ)[\-‐]?\s*[0-9]"
)

# 「同じ行」とみなすY座標の許容誤差。当初1.0(ほぼ完全一致)だったが、
# WA1-1B-02Aで、同じ行のはずのマーク+本数セル(Y=10984.3)と設計符号+
# サイズセル(Y=10978.0)の間に6.3程度のズレがあり、行として認識されず
# 2マーク(WA11-1TB441-5/6)が丸ごと欠落する不具合があった(2026-09-14)。
# 行間隔は実測240前後あるため、10.0に広げても別の行と混同するリスクは
# 十分小さい。
SAME_ROW_Y_TOLERANCE = 10.0


def _looks_like_size(text: str) -> bool:
    return bool(SIZE_PATTERN.match(text.strip()))


# ---------------------------------------------------------------------------
# 製品情報テーブルの検出
# ---------------------------------------------------------------------------

@dataclass
class ProductRow:
    mark: str
    design_code: str | None
    size: str | None
    count: str | None
    weight: str | None
    x_mark: float
    x_next: float  # 本数セルのX座標(長さ判定のX範囲右端に使う)
    y: float
    length: float | None = None
    left_joint: str | None = None
    right_joint: str | None = None
    drawing_number: str | None = None


def find_product_rows(tdf: tb.TdfData) -> list[ProductRow]:
    """「N台」セルを起点に、同じ行の製品マーク・設計符号・サイズを逆算して集める。

    本体・参照のどちらも行の起点になり得るため、resolve_text()で文字列を解決した
    上で、実際の配置座標(重複しない)ごとに1行として扱う。
    """
    dai_cells = []
    for r in tdf.texts:
        if r.rot != 0:
            continue
        text = tdf.resolve_text(r)
        if text and tb.is_dai_cell(text):
            dai_cells.append(r)

    rows: list[ProductRow] = []
    for dai in dai_cells:
        # 同じ行とみなせるY範囲で、Xがdaiより小さい通常テキスト(rot=0)を
        # 集めて左から並べる
        same_row = [
            r for r in tdf.texts
            if r.rot == 0 and abs(r.y - dai.y) < SAME_ROW_Y_TOLERANCE and r.x < dai.x and tdf.resolve_text(r)
        ]
        same_row.sort(key=lambda r: r.x)
        if len(same_row) < 3:
            continue
        # 直近3つを [製品マーク, 設計符号, サイズ] の順とみなす(サイズがdaiに一番近い)
        size_rec, code_rec, mark_rec = same_row[-1], same_row[-2], same_row[-3]
        size_text = tdf.resolve_text(size_rec)
        if not _looks_like_size(size_text):
            continue

        # 重量(daiより右、ほぼ同じY、"N.Nkg"形式)があれば拾う。重量セルは本数セルと
        # 数ピクセル分Yがずれることがあるため、許容幅を広めに取る。
        weight_text = None
        same_row_right = [
            r for r in tdf.texts
            if r.rot == 0 and abs(r.y - dai.y) < 100.0 and r.x > dai.x and tdf.resolve_text(r)
        ]
        same_row_right.sort(key=lambda r: r.x)
        if same_row_right:
            cand = tdf.resolve_text(same_row_right[0])
            if cand and cand.rstrip().endswith("kg"):
                weight_text = cand

        rows.append(ProductRow(
            mark=tdf.resolve_text(mark_rec), design_code=tdf.resolve_text(code_rec),
            size=size_text, count=tdf.resolve_text(dai), weight=weight_text,
            x_mark=mark_rec.x, x_next=dai.x, y=dai.y,
        ))

    rows.sort(key=lambda r: r.x_mark)
    return rows


_DRAWING_NUMBER_PATTERN = re.compile(r"^[A-Za-z][A-Za-z0-9\-]{2,19}$")


def _find_zumenbango_label_pos(tdf: tb.TdfData) -> tuple[float, float] | None:
    """「図」「面」「番」「号」の4文字が同じY・等間隔でX方向に連続して並ぶ
    箇所(タイトル欄の「図面番号」ラベルそのもの)を探し、その(Y, X中心)を
    返す。単独の「図」「面」等の文字が他の場所にも現れることがあるため、
    4文字が実際に連続しているものだけを対象にする。"""
    chars: dict[str, list] = {"図": [], "面": [], "番": [], "号": []}
    for r in tdf.texts:
        text = tdf.resolve_text(r)
        if text and r.rot == 0 and text.strip() in chars:
            chars[text.strip()].append(r)

    if not all(chars.values()):
        return None

    for zu in chars["図"]:
        row = [zu]
        cur = zu
        for nextchar in ("面", "番", "号"):
            cand = [
                r for r in chars[nextchar]
                if abs(r.y - zu.y) < 1.0 and cur.x < r.x < cur.x + cur.h * 3
            ]
            if not cand:
                break
            cur = min(cand, key=lambda r: r.x)
            row.append(cur)
        else:
            return (zu.y, sum(r.x for r in row) / len(row))
    return None


def find_drawing_number(tdf: tb.TdfData) -> str | None:
    """タイトル欄の図番らしきテキストを推定する。

    まず「図面番号」ラベル(4文字が連続して並ぶ箇所)を探し、その近く
    (Y座標の差が500mm以内)にある英数字+ハイフンの候補の中で、ラベルに
    最もX方向で近いものを図番として採用する。ラベルが見つからない場合や
    近くに候補が無い場合は、従来通り文字高さが最も大きいものを採用する
    (フォールバック)。

    2026-09-14、鳥取現場図面EA1-1G-14で、製品テーブルの「設計符号」欄が
    たまたまタイトル欄の図番と同じ文字高さだったため誤って設計符号の方を
    採用してしまう問題が発覚し、位置情報(ラベルとの近さ)を優先する方式に
    改訂した。"""
    candidates = []
    for r in tdf.texts:
        if r.text is None or r.rot != 0:
            continue
        t = r.text.strip()
        if _DRAWING_NUMBER_PATTERN.match(t) and not t.isdigit():
            candidates.append(r)
    if not candidates:
        return None

    label_pos = _find_zumenbango_label_pos(tdf)
    if label_pos is not None:
        label_y, label_x = label_pos
        near = [r for r in candidates if abs(r.y - label_y) < 500.0]
        if near:
            near.sort(key=lambda r: abs(r.x - label_x))
            return near[0].text.strip()

    candidates.sort(key=lambda r: -r.h)
    return candidates[0].text.strip()


# ---------------------------------------------------------------------------
# 「長さ」の判定
# ---------------------------------------------------------------------------

def _axis_key(rot_deg: float, xr: float, yr: float) -> tuple[str, float]:
    r = rot_deg % 180
    if r > 90:
        r = 180 - r
    if r < 45:
        return ("yr", round(yr, 0))
    return ("xr", round(xr, 0))


def determine_length(tdf: tb.TdfData, x_min: float, x_max: float, y_ref: float,
                      min_value: float = 300.0, y_max: float | None = None) -> tuple[float | None, dict]:
    """製品のX範囲・Y範囲内から「長さ」を1つに絞り込む。

    `y_max`を指定すると、Y範囲の上限をそこで打ち切る(1枚の図面に梁の
    立面が2段に描かれている場合、下段の行から見て上段の行より上は
    別の梁の寸法線なので含めてはいけない。2026-09-14追加)。指定が無い
    場合は従来通り上方向は無制限。

    戻り値: (長さの値 or None, デバッグ情報dict)
    """
    seen_coords = set()
    items = []  # (text, value, x, y, rot, xr, yr)
    for r in tdf.texts:
        text = tdf.resolve_text(r)
        v = _try_float(text)
        if v is None:
            continue
        if not (x_min <= r.x <= x_max and r.y > y_ref and (y_max is None or r.y < y_max)):
            continue
        key = (round(r.x, 1), round(r.y, 1))
        if key in seen_coords:
            continue
        seen_coords.add(key)
        rot = tdf.resolve_rot(r)
        theta = -rot
        xr = r.x * math.cos(theta) - r.y * math.sin(theta)
        yr = r.x * math.sin(theta) + r.y * math.cos(theta)
        items.append((text, v, r.x, r.y, rot, xr, yr))

    if not items:
        return None, {"reason": "no numeric texts in range"}

    # 主軸方向(最頻のrot)を推定。回転行列の計算に使うため、符号を保持した
    # 生の角度(-180〜180)のまま集計する(% 180で丸めると符号情報が失われ、
    # 後段の回転変換が180度近くずれてしまう)。
    # 1/5フィルタ(次のステップ)より前の、絞り込み前の全テキストから求める。
    # 先に1/5フィルタをかけてしまうと、たまたま小さい値ばかりが多い向きの
    # テキストが削られて、残った少数派の向きに主軸が引っ張られることがある
    # (2026-09-14、夢前現場図面WA1-1G-02で確認。フィルタ順序を修正)。
    rot_counts: dict[float, int] = {}
    for _t, _v, _x, _y, rot, _xr, _yr in items:
        key = round(math.degrees(rot), 1)
        rot_counts[key] = rot_counts.get(key, 0) + 1
    main_axis_deg = max(rot_counts.items(), key=lambda kv: kv[1])[0]

    # 範囲内の最大値の1/5未満の値は「取付ピッチ」等の小さい寸法とみなし、
    # 長さ候補から除外する(小さすぎる値がスコアリングで誤って勝つのを防ぐ)。
    max_v = max(v for _t, v, *_ in items)
    threshold = max_v / 5.0
    items = [it for it in items if it[1] >= threshold]
    if not items:
        return None, {"reason": "no items after 1/5 filter"}

    # 出現回数
    occurrence: dict[float, list] = {}
    for text, v, x, y, rot, xr, yr in items:
        occurrence.setdefault(round(v, 2), []).append((x, y))

    # 回転後座標でのグループ化(取付ピッチ系列の検出)
    group_map: dict[tuple, list] = {}
    for text, v, x, y, rot, xr, yr in items:
        key = _axis_key(math.degrees(rot), xr, yr)
        group_map.setdefault(key, []).append((text, v, x, y, xr, yr))

    excluded = set()
    for val, coords in occurrence.items():
        if len(coords) >= 5:
            excluded.add(val)
    for key, members in group_map.items():
        if len(members) >= 3:
            for text, v, x, y, xr, yr in members:
                excluded.add(round(v, 2))

    group_sums = []
    for key, members in group_map.items():
        total = sum(v for text, v, x, y, xr, yr in members)
        group_sums.append((key, total, len(members), [m[0] for m in members]))

    # 除外された値でも他系列合計と一致するなら復活
    revived = set()
    for val in list(excluded):
        for key, total, cnt, texts in group_sums:
            member_vals = {round(_try_float(tx), 2) for tx in texts if _try_float(tx) is not None}
            if val in member_vals:
                continue
            if abs(total - val) < 1.0:
                revived.add(val)
                break

    def line_angle_deg(x1, y1, x2, y2):
        return math.degrees(math.atan2(y2 - y1, x2 - x1)) % 180

    def has_primary_line(val, angle_tol=5.0):
        """ref0==3(実寸法師が自動生成する主要な寸法線)の一致直線が
        1本でもあれば、取付ピッチ等の繰り返し寸法とみなす除外ロジックを
        無視して良いという確実な証拠になる(2026-09-14、EA2-1G-11で、本来
        の長さが他の数値と同じ行に並んでいたために誤って除外されていた
        問題を受けて追加)。"""
        for ln in tdf.lines:
            if ln.ref0 != 3:
                continue
            mid_x = (ln.x1 + ln.x2) / 2
            if not (x_min <= mid_x <= x_max and ln.y1 > y_ref and ln.y2 > y_ref
                    and (y_max is None or (ln.y1 < y_max and ln.y2 < y_max))):
                continue
            if abs(ln.length - val) >= 0.5:
                continue
            ang = line_angle_deg(ln.x1, ln.y1, ln.x2, ln.y2)
            diff = abs(ang - main_axis_deg)
            diff = min(diff, 180 - diff)
            if diff <= angle_tol:
                return True
        return False

    candidates = [
        round(v, 2) for v in occurrence
        if v > min_value and (
            round(v, 2) not in excluded
            or round(v, 2) in revived
            or has_primary_line(round(v, 2))
        )
    ]
    if not candidates:
        return None, {"reason": "no candidates after filtering", "main_axis_deg": main_axis_deg}

    def count_matching_lines(val, angle_tol=5.0):
        cnt = 0
        primary_cnt = 0  # ref0==3(実寸法師が自動生成する主要な寸法線)の本数
        for ln in tdf.lines:
            mid_x = (ln.x1 + ln.x2) / 2
            if not (x_min <= mid_x <= x_max and ln.y1 > y_ref and ln.y2 > y_ref
                    and (y_max is None or (ln.y1 < y_max and ln.y2 < y_max))):
                continue
            if abs(ln.length - val) >= 0.5:
                continue
            ang = line_angle_deg(ln.x1, ln.y1, ln.x2, ln.y2)
            diff = abs(ang - main_axis_deg)
            diff = min(diff, 180 - diff)
            if diff <= angle_tol:
                cnt += 1
                if ln.ref0 == 3:
                    primary_cnt += 1
        return cnt, primary_cnt

    scored = []
    for val in candidates:
        a, primary_cnt = count_matching_lines(val)
        b = sum(1 for key, total, cnt, texts in group_sums if abs(total - val) < 1.0)
        scored.append((val, a, b, primary_cnt))

    # ref0==3の本数(主要寸法線としての一致数)を最優先。同点ならa、さらに同点ならb。
    # ボルトピッチ・リブ等の手動追加線はref0が3以外の値を持つため、これらの
    # 候補は自然と後順位になる(2026-09-14、鳥取現場図面での検証で確立)。
    scored.sort(key=lambda x: (-x[3], -x[1], -x[2]))

    best = scored[0]
    debug = {"main_axis_deg": main_axis_deg, "candidates": scored,
              "threshold": threshold, "max_v": max_v}
    if best[3] == 0 and best[1] == 0:
        # フォールバック(2026-09-14追加): 小梁の斜め材で、寸法テキスト自身の
        # rotが部材の傾きと一致しないケース(EA2-1B-03のEA21-1b2)があり、
        # main_axis_deg(テキストのrot多数決)が誤って0度になってしまい、
        # 本来の長さ(斜め方向の直線)が「主軸不一致」で除外されてしまっていた。
        # この場合に限り、向きを問わず「長さが完全一致する直線の本数」だけで
        # 再スコアリングする(ユーザー確認: EA21-1b2-1/2の正解2688.6は、
        # 向きが斜め[157.17度]の直線が範囲内に7本[最多]あることで裏付けられる)。
        # 通常ケース(大梁48ファイル・小梁の他の行)は最初のスコアリングで
        # 候補が見つかるため、このフォールバックが実行されることはない。
        any_angle_scored = []
        for val, _a, b, _primary_cnt in scored:
            cnt = 0
            primary_cnt = 0
            for ln in tdf.lines:
                mid_x = (ln.x1 + ln.x2) / 2
                if not (x_min <= mid_x <= x_max and ln.y1 > y_ref and ln.y2 > y_ref
                        and (y_max is None or (ln.y1 < y_max and ln.y2 < y_max))):
                    continue
                if abs(ln.length - val) >= 0.5:
                    continue
                cnt += 1
                if ln.ref0 == 3:
                    primary_cnt += 1
            any_angle_scored.append((val, cnt, b, primary_cnt))
        any_angle_scored.sort(key=lambda x: (-x[3], -x[1], -x[2]))
        fallback_best = any_angle_scored[0]
        debug["any_angle_candidates"] = any_angle_scored
        if fallback_best[3] == 0 and fallback_best[1] == 0:
            return None, {**debug, "reason": "no direction-matched line for any candidate"}

        # 採用した値に実際に一致した直線自身の角度を、真の主軸として
        # 採用し直す(2026-09-14追加。フォールバック前のmain_axis_degは
        # テキストのrot多数決による誤った値[0度]のままだったため、
        # 「種別」列[普通梁/斜め梁]の判定にも誤って使われ、EA21-1b2
        # [実際は斜め梁]が「普通梁」と誤表示される不具合があった)。
        matched_angles: dict[float, int] = {}
        for ln in tdf.lines:
            mid_x = (ln.x1 + ln.x2) / 2
            if not (x_min <= mid_x <= x_max and ln.y1 > y_ref and ln.y2 > y_ref
                    and (y_max is None or (ln.y1 < y_max and ln.y2 < y_max))):
                continue
            if abs(ln.length - fallback_best[0]) >= 0.5:
                continue
            ang = round(line_angle_deg(ln.x1, ln.y1, ln.x2, ln.y2), 1)
            matched_angles[ang] = matched_angles.get(ang, 0) + 1
        if matched_angles:
            debug["main_axis_deg"] = max(matched_angles.items(), key=lambda kv: kv[1])[0]

        return fallback_best[0], {**debug, "reason": "axis-less fallback used"}
    return best[0], debug


# ---------------------------------------------------------------------------
# 「左継手・右継手」の判定
# ---------------------------------------------------------------------------

def find_stadium_centers(tdf: tb.TdfData, angle_tol_deg: float = 1.0,
                          radius_tol: float = 0.5, max_dist_factor: float = 20.0):
    """長丸(半径一致・開始角180度差の円弧ペア)の中心座標を全部求める。"""
    arcs = [a for a in tdf.arcs if a.a1 is not None]
    centers = []
    used = set()
    n = len(arcs)
    for i in range(n):
        a = arcs[i]
        if a.offset in used:
            continue
        for j in range(i + 1, n):
            b = arcs[j]
            if b.offset in used:
                continue
            if abs(a.r - b.r) > radius_tol:
                continue
            a1d = math.degrees(a.a1)
            b1d = math.degrees(b.a1)
            diff = abs((a1d - b1d) % 360)
            diff = min(diff, 360 - diff)
            if abs(diff - 180.0) > angle_tol_deg:
                continue
            dist = math.hypot(b.cx - a.cx, b.cy - a.cy)
            if not (0.1 < dist < a.r * max_dist_factor):
                continue
            mid_x = (a.cx + b.cx) / 2
            mid_y = (a.cy + b.cy) / 2
            centers.append((mid_x, mid_y, a.r))
            used.add(a.offset)
            used.add(b.offset)
            break
    return centers


JOINT_CODE_PATTERN = re.compile(r"^GJ")


def determine_joints(tdf: tb.TdfData, x_min: float, x_max: float, y_ref: float,
                      main_axis_deg: float, length_value: float,
                      x_tolerance_factor: float = 0.6,
                      y_max: float | None = None) -> tuple[str | None, str | None]:
    """長丸で囲まれた継手候補のうち、継手コード特有の文字列パターン(`GJ`で
    始まる)に一致するものだけを対象に、製品マーク側(左)・本数セル側(右)の
    それぞれに最も近いものを左右継手として採用する。

    継手コードは図面内で必ず`GJ`始まりであり(`P441`や`TB888`等の無関係な
    部材記号はこのパターンに一致しない)、「N台」パターンと同様に文字列
    そのもので確実に識別できる(2026-09-14、鳥取現場図面での検証で確立)。
    片側にしか継手が無い図面では、該当側の候補が遠すぎる(x_tolerance_factor
    ×lengthを超える)場合にNoneを返す。

    `y_max`を指定すると、長さ判定(determine_length)と同様にY範囲の上限を
    打ち切る(1枚の図面に梁の立面が2段に描かれている場合、下段の行から
    見て上段の行より上にある継手は別の梁のものなので含めてはいけない。
    2026-09-14追加)。
    """
    centers = find_stadium_centers(tdf)
    theta = -math.radians(main_axis_deg)

    def rotate(x, y):
        xr = x * math.cos(theta) - y * math.sin(theta)
        yr = x * math.sin(theta) + y * math.cos(theta)
        return xr, yr

    x_mark_r, _ = rotate(x_min, y_ref)
    x_next_r, _ = rotate(x_max, y_ref)

    joint_candidates = []
    for mx, my, r in centers:
        near = [rec for rec in tdf.texts if abs(rec.x - mx) < 3 and abs(rec.y - my) < 3]
        for rec in near:
            text = tdf.resolve_text(rec)
            if not text or not JOINT_CODE_PATTERN.match(text):
                continue
            if my <= y_ref or (y_max is not None and my >= y_max):
                continue
            xr, yr = rotate(mx, my)
            joint_candidates.append((text, xr))

    if not joint_candidates:
        return None, None

    threshold = length_value * x_tolerance_factor
    left_best = min(joint_candidates, key=lambda c: abs(c[1] - x_mark_r))
    right_best = min(joint_candidates, key=lambda c: abs(c[1] - x_next_r))
    left = left_best[0] if abs(left_best[1] - x_mark_r) <= threshold else None
    right = right_best[0] if abs(right_best[1] - x_next_r) <= threshold else None
    return left, right


# ---------------------------------------------------------------------------
# メイン抽出処理
# ---------------------------------------------------------------------------

RELAX_MARGIN = 5000.0


def extract(path: str) -> list[ProductRow]:
    tdf = tb.load(path)
    drawing_number = find_drawing_number(tdf)
    rows = find_product_rows(tdf)

    for i, row in enumerate(rows):
        if len(rows) == 1:
            # 製品が1種類だけの図面は、他の製品との取り違えリスクが無いため
            # X範囲を左右5000mmずつ緩和し、長さ候補の見落としを減らす。
            x_min = row.x_mark - RELAX_MARGIN
            x_max = row.x_next + RELAX_MARGIN
        else:
            x_min = row.x_mark
            x_max = row.x_next
        y_ref = row.y
        length_value, debug = determine_length(tdf, x_min, x_max, y_ref)
        row.length = length_value
        if length_value is not None:
            main_axis_deg = debug.get("main_axis_deg", 0.0)
            # 継手判定は緩和前の製品自身のマーク/本数セル座標を基準にする
            # (緩和後のx_min/x_maxを使うと、判定基準点が製品本体から
            # 大きくずれて別の候補を誤って拾ってしまう)。
            left, right = determine_joints(
                tdf, row.x_mark, row.x_next, y_ref, main_axis_deg, length_value
            )
            row.left_joint = left
            row.right_joint = right
        row.drawing_number = drawing_number

    return rows


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 2:
        print("使い方: python tdf_master_extractor.py <tdfパス>")
        raise SystemExit(1)

    rows = extract(sys.argv[1])
    dn = find_drawing_number(tb.load(sys.argv[1]))
    print(f"図番: {dn}")
    for r in rows:
        print(f"  製品マーク={r.mark} 設計符号={r.design_code} サイズ={r.size} "
              f"本数={r.count} 重量={r.weight} 長さ={r.length} "
              f"左継手={r.left_joint} 右継手={r.right_joint}")
