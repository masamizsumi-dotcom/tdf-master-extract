"""実寸法師(.tdf)から製品マスタ情報を抽出するロジックのうち、「1製品に複数
マークが並び、設計符号・サイズを共有する」パターン専用の検出コード。

2026-09-14、小梁(1B系)図面の実機解析で確立。既存の`tdf_master_extractor.py`
(「1マーク=1製品」の大梁パターン、EA1-1G系等で検証済み)には一切手を
加えず、こちらは完全に別コードとして追加する。1つのファイル内に両方の
パターンが混在することがあるため(例: EA1-1B-06)、呼び出し側で両方の
関数を実行して結果をマージすることを想定している。

前提となる図面構造:
  - マークX座標・本数(N台)セルのX座標が完全に一致する行同士を、1つの
    共有テーブルとみなす(隣接する別テーブルはX座標が異なるため混ざらない)。
  - 設計符号・サイズはテーブル内で1回だけ描画され、そのY座標は
    「テーブル内の最大Y・最小Yの中間点((最大Y+最小Y)/2)」に一致する
    (2026-09-14、EA1-1B-01[2行]・EA1-1B-02[3行]・EA1-1B-06[4行]の
    3パターンで検証: いずれも計算値と実測値が完全一致)。
  - ×印で削除された行は、マークセル・本数セルの矩形範囲に対角線が交差する
    斜め線ペア(同じ矩形の対角線同士)が重なっている。これを検出して
    除外する(2026-09-14、EA1-1B-06の-7行[ref0=220069891の斜め線2組]で確認)。
"""
from __future__ import annotations

import math
import re

import tdf_master_extractor as ex
import tdf_binary as tb

_MARK_NUM_PATTERN = re.compile(r"-(\d+)([A-Za-z]?)$")


def mark_sort_key(mark_text: str | None):
    """マーク文字列末尾の「-番号」を数値として昇順に並べるためのキー
    (2026-09-14、共有テーブル内の並び順がY座標の描画順になってしまい、
    -2の次に-1が来る逆順になる不具合を修正)。呼び出し側(extract_to_excel.py
    等)でも、ex.find_product_rows()由来の行とexm由来の行が混在する行を
    最終的に並べ直す際にこのキーを使うこと(片方の関数だけで昇順にしても、
    もう片方の関数が見つけた同じグループの行と混ざると順序が崩れるため)。"""
    if not mark_text:
        return (mark_text or "",)
    m = _MARK_NUM_PATTERN.search(mark_text)
    if not m:
        return (mark_text,)
    return (int(m.group(1)), m.group(2))


def mark_prefix(mark_text: str | None) -> str:
    """マーク文字列から末尾の「-番号」を除いた接頭辞を返す
    (例: "EA12-1B34-4"→"EA12-1B34")。同じ接頭辞を持つマークは、物理的に
    どの表(列・段)に描かれているかに関係なく1つの製品とみなし、番号の
    昇順で並べるために使う(2026-09-14、EA1-1B-06でB34マークが複数の表に
    分かれて描かれていても、全体としては番号順に並べたいという指示)。"""
    if not mark_text:
        return mark_text or ""
    m = _MARK_NUM_PATTERN.search(mark_text)
    if not m:
        return mark_text
    return mark_text[:m.start()]


_AXIS_ALIGN_TOL_DEG = 5.0
_XMARK_CORNER_TOL = 2.0
_SHARED_CODE_Y_TOL = 2.0

# 継手候補の探索閾値の最低保証値。実測データ(95ファイル)で、短い製品
# (300〜1500mm程度)の実在する継手候補までの距離が最大2908.5だったため、
# 余裕を見て3000を採用する(2026-09-14)。
JOINT_MIN_THRESHOLD = 3000.0


def _is_axis_aligned(ln: tb.LineRecord) -> bool:
    ang = math.degrees(math.atan2(ln.y2 - ln.y1, ln.x2 - ln.x1)) % 180
    return ang < _AXIS_ALIGN_TOL_DEG or ang > 180 - _AXIS_ALIGN_TOL_DEG or abs(ang - 90) < _AXIS_ALIGN_TOL_DEG


def _close(p: tuple[float, float], q: tuple[float, float], tol: float) -> bool:
    return abs(p[0] - q[0]) < tol and abs(p[1] - q[1]) < tol


def _forms_x_cross(a: tb.LineRecord, b: tb.LineRecord, tol: float = _XMARK_CORNER_TOL):
    """2本の斜め線が、同じ矩形の対角線同士(=X字)を成すかどうか。"""
    corner1 = (a.x1, a.y2)
    corner2 = (a.x2, a.y1)
    b_pts = [(b.x1, b.y1), (b.x2, b.y2)]
    if (_close(b_pts[0], corner1, tol) and _close(b_pts[1], corner2, tol)) or \
       (_close(b_pts[0], corner2, tol) and _close(b_pts[1], corner1, tol)):
        x0, x1 = sorted((a.x1, a.x2))
        y0, y1 = sorted((a.y1, a.y2))
        return (x0, x1, y0, y1)
    return None


def find_x_marks(tdf: tb.TdfData) -> list[tuple[float, float, float, float]]:
    """ファイル内の全ての×印(削除マーク)の矩形範囲(x0,x1,y0,y1)を求める。"""
    diag = [ln for ln in tdf.lines if not _is_axis_aligned(ln)]
    boxes = []
    used = set()
    n = len(diag)
    for i in range(n):
        if diag[i].offset in used:
            continue
        for j in range(i + 1, n):
            if diag[j].offset in used:
                continue
            box = _forms_x_cross(diag[i], diag[j])
            if box is not None:
                boxes.append(box)
                used.add(diag[i].offset)
                used.add(diag[j].offset)
                break
    return boxes


def _point_in_any_box(x: float, y: float, boxes: list[tuple[float, float, float, float]]) -> bool:
    for x0, x1, y0, y1 in boxes:
        if x0 <= x <= x1 and y0 <= y <= y1:
            return True
    return False


def filter_deleted_rows(tdf: tb.TdfData, rows: list) -> list:
    """マーク・本数セルのどちらかが×印(削除マーク)の矩形範囲内にある行を
    除外する。`find_product_rows_shared_group`は自分が拾った行に対しては
    既にこのチェックを行っているが、`tdf_master_extractor.find_product_rows`
    (1マーク=1製品パターン、設計符号・サイズが自分の行にインラインで
    描かれているケース)には削除マークのチェックが元々無いため、行の
    出どころに関わらず全行に対してここでまとめてチェックする
    (2026-09-14、EA1-1B-08のEA12-1TB489-2[×印済み]がex.find_product_rows
    経由で素通りしていた不具合を修正)。"""
    x_marks = find_x_marks(tdf)
    if not x_marks:
        return rows
    return [
        r for r in rows
        if not (_point_in_any_box(r.x_mark, r.y, x_marks) or _point_in_any_box(r.x_next, r.y, x_marks))
    ]


def _find_shared_code_size(tdf: tb.TdfData, x_min: float, x_max: float, y_lo: float, y_hi: float):
    """テーブルの行のY範囲[y_lo, y_hi]内にある共有設計符号+サイズを探す。

    2026-09-14当初は「(最大Y+最小Y)/2」の1点だけを探していたが、EA2-1B-03
    (阪和ではなく鳥取EA2現場)の列1(-1,-2,-5,-6)で、ラベルが4行全体の中間点
    ではなく先頭2行(-1,-2)の中間点にだけ描かれ、-5/-6専用のラベルは存在しない
    (同じ設計として省略されている)ケースが見つかった。中間点1点への一致
    ではなくグループの行が実際に描かれているY範囲全体を探索するよう緩和する
    (範囲を絞るx_min/x_maxは元々そのグループ専用の列幅のため、他グループの
    ラベルを誤って拾うリスクは低い)。"""
    same_line = [
        r for r in tdf.texts
        if r.rot == 0 and x_min < r.x < x_max
        and y_lo - _SHARED_CODE_Y_TOL <= r.y <= y_hi + _SHARED_CODE_Y_TOL
        and tdf.resolve_text(r)
    ]
    same_line.sort(key=lambda r: r.x)
    if len(same_line) < 2:
        return None, None
    size_rec, code_rec = same_line[-1], same_line[-2]
    size_text = tdf.resolve_text(size_rec)
    if not ex._looks_like_size(size_text):
        return None, None
    return tdf.resolve_text(code_rec), size_text


def find_product_rows_shared_group(
    tdf: tb.TdfData, existing_positions: set[tuple[float, float]] | None = None,
) -> list[ex.ProductRow]:
    """「マーク+本数」だけが行ごとに並び、設計符号・サイズを複数行で共有する
    テーブルパターンを検出する。`existing_positions`には
    `tdf_master_extractor.find_product_rows()`で既に拾われたdaiセルの
    (round(x,1), round(y,1))座標を渡すと、二重集計を避けられる。
    """
    existing_positions = existing_positions or set()
    x_marks = find_x_marks(tdf)

    dai_cells = []
    for r in tdf.texts:
        if r.rot != 0:
            continue
        text = tdf.resolve_text(r)
        if text and tb.is_dai_cell(text):
            key = (round(r.x, 1), round(r.y, 1))
            if key in existing_positions:
                continue
            dai_cells.append(r)

    # 各daiセルについて、同じ行(ex.SAME_ROW_Y_TOLERANCE・X<dai.x)のテキストを
    # 集める。3つ以上あり末尾がサイズらしい場合は「1マーク=1製品」パターンで
    # 既に処理済みのはずなので、ここでは対象にしない(念のためのガード)。
    candidates = []
    for dai in dai_cells:
        same_row = [
            r for r in tdf.texts
            if r.rot == 0 and abs(r.y - dai.y) < ex.SAME_ROW_Y_TOLERANCE and r.x < dai.x and tdf.resolve_text(r)
        ]
        same_row.sort(key=lambda r: r.x)
        if len(same_row) >= 3 and ex._looks_like_size(tdf.resolve_text(same_row[-1])):
            continue
        if not same_row:
            continue
        mark_rec = same_row[-1]
        candidates.append((mark_rec, dai))

    # マークX・本数Xがほぼぴったり一致する行同士を同一テーブルとしてグループ化
    #
    # 注意(2026-09-14): 同じマーク文字列が2箇所に出現することがある
    # (例: EA2-1B-03のEA21-1B35A-5/6は、列1[元の位置]と列3[詳細図の凡例
    # テーブル]の両方に描かれている)。これは「列1の記載を列3に移設した」
    # という設計変更を示す表現で、移設元(列1側)には×印(削除マーク)が
    # 描かれ、移設先(列3側)には描かれていない。ここで文字列の重複を
    # 理由に一方を機械的に除外することはしない。×印が付いている側は
    # 後段の`filter_deleted_rows`/このグループ内の削除チェックで自然に
    # 除外され、×印が無い側(=現在有効な記載)だけが残る。KEY PLAN上の
    # 位置マーカーとの重複(TB577・b2等、どちらも×印は無い)でも、
    # 本体/参照どちらの座標がテーブルの実データとして扱われるべきかは
    # 常に「dai(本数)セルと正しく対になっているか」で決まるため、
    # 文字列の重複だけで除外する判断はしない方が安全。
    groups: dict[tuple[float, float], list[tuple]] = {}
    for mark_rec, dai in candidates:
        key = (round(mark_rec.x, 1), round(dai.x, 1))
        groups.setdefault(key, []).append((mark_rec, dai))

    rows: list[ex.ProductRow] = []
    for _key, members in groups.items():
        # 同じ(マークX,本数X)でも、繰り返しテンプレートにより全く別の
        # フレーム(離れたY位置)にある表が同じ列位置を再利用していることが
        # ある(2026-09-14、EA2-1B-06で、Y3位置のフレーム[-3,-4,-5,-10,
        # -11,-12]とY8位置のフレーム[-17,-18]が同じ列位置になり、1つの
        # 巨大なグループ[Y範囲8500mm超]として誤って統合されて共有ラベルの
        # 検索に失敗し、8マークが丸ごと欠落していた不具合)。Yが大きく
        # 離れている(目安3000mm以上、通常の行間隔240前後に対して十分大きい
        # 閾値)メンバー同士は別の表とみなして分割し、それぞれ独立に
        # 共有設計符号・サイズを探す。
        members_by_y = sorted(members, key=lambda item: item[1].y)
        clusters: list[list[tuple]] = [[members_by_y[0]]]
        for item in members_by_y[1:]:
            if item[1].y - clusters[-1][-1][1].y > 3000.0:
                clusters.append([item])
            else:
                clusters[-1].append(item)

        for cluster in clusters:
            # クラスタキーはグルーピング用に丸めた値なので、検索範囲には
            # 必ず実際の(丸めていない)座標を使う。丸めた値をそのまま使うと、
            # ごくわずかな丸め誤差で本数セル自身が検索範囲に入り込み、
            # 「N台」が末尾(サイズの位置)に来て設計符号・サイズの検出に
            # 失敗することがある(2026-09-14、EA1-1B-08のEA12-1TB489-1が
            # この理由で丸ごと欠落していた不具合を修正)。
            mark_x = cluster[0][0].x
            dai_x = cluster[0][1].x
            ys = [dai.y for _mark, dai in cluster]
            design_code, size_text = _find_shared_code_size(tdf, mark_x, dai_x, min(ys), max(ys))
            if design_code is None:
                continue  # 共有セルが見つからない=未知のパターン。既存挙動どおり無視する

            cluster_sorted = sorted(cluster, key=lambda item: mark_sort_key(tdf.resolve_text(item[0])))
            for mark_rec, dai in cluster_sorted:
                if _point_in_any_box(mark_rec.x, mark_rec.y, x_marks) or _point_in_any_box(dai.x, dai.y, x_marks):
                    continue  # ×印で削除された行

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

                rows.append(ex.ProductRow(
                    mark=tdf.resolve_text(mark_rec), design_code=design_code, size=size_text,
                    count=tdf.resolve_text(dai), weight=weight_text,
                    x_mark=mark_rec.x, x_next=dai.x, y=dai.y,
                ))

    rows.sort(key=lambda r: r.x_mark)
    return rows


# ---------------------------------------------------------------------------
# 「継手」の判定(小梁系専用: プレフィックスに頼らない版)
# ---------------------------------------------------------------------------

def _find_reference_line_endpoints(tdf: tb.TdfData, x_min: float, x_max: float, y_ref: float,
                                    y_max: float | None, length_value: float):
    """`determine_length`が採用した長さと一致する直線のうち、行の`y_ref`に
    最も近いものを実際の部材線とみなし、その左右端点の(x, y)座標を返す。

    同じ長さの直線が(別の位置の取付ピッチ・偶然の一致等で)複数見つかる
    ことがあるため、単純な最初の一致ではなくY方向の近さで選ぶ
    (2026-09-14、WA1-1B-03のEA21-1TB788-1/2/3で、無関係な遠いY位置に
    たまたま同じ長さの直線[取付ピッチ等の繰り返し]があり、そのX位置に
    近いという理由だけで別グループの継手候補[TB441]を誤って右継手として
    採用してしまっていた不具合を受けて追加。継手候補は部材の実際の端点の
    近くに描かれるという前提に立ち、以降の距離判定はX方向だけでなく
    実際の端点からの2次元距離で行う)。

    一致する直線が見つからない場合はNoneを返す(呼び出し側はマーク/本数
    セル自身の座標を使う既存の方式にフォールバックする)。
    """
    best = None
    best_dist = None
    for ln in tdf.lines:
        mid_x = (ln.x1 + ln.x2) / 2
        if not (x_min <= mid_x <= x_max and ln.y1 > y_ref and ln.y2 > y_ref
                and (y_max is None or (ln.y1 < y_max and ln.y2 < y_max))):
            continue
        if abs(ln.length - length_value) >= 0.5:
            continue
        mid_y = (ln.y1 + ln.y2) / 2
        dist = abs(mid_y - y_ref)
        if best_dist is None or dist < best_dist:
            best_dist = dist
            best = ln
    if best is None:
        return None
    if best.x1 <= best.x2:
        return (best.x1, best.y1), (best.x2, best.y2)
    return (best.x2, best.y2), (best.x1, best.y1)


def assign_joints_batch(tdf: tb.TdfData, rows: list, tier_info: dict, lengths: dict) -> dict:
    """長丸(スタジアム形状)に囲まれたテキストは全て継手マークの一種である、
    という前提に立つ(2026-09-14、ユーザー確認: 大梁の`GJ`のような専用の
    文字列プレフィックスは無く、小梁では`TB441`や`TB489`等、設計符号と
    同じ文字列がそのまま継手マークとしても使われる)。GJプレフィックスの
    代わりに位置(距離・段のY範囲)だけで判定する。

    ファイル内の行をまとめて渡し、「同じグループ(=同じ(x_mark,x_next)を
    持つ行の集まり。共有テーブルなら複数行、単独行なら1行)」単位で
    継手候補を割り当てる。1つの継手候補は1グループにしか属さない
    (早い者勝ちではなく、複数グループが同じ候補を欲しがった場合は本当に
    距離が近い方が獲得する)。グループ内では候補を共有する(共有テーブルの
    継手は1つしか描かれないため)。

    2026-09-14、EA1-1B-09のEA12-1B34-51(独立グループ)が、実際には別の
    グループ(EA12-1TB489-2)によりふさわしい継手候補をX距離だけで見て
    誤って奪ってしまう不具合と、1グループが複数の候補で「勝利」した際に
    距離を無視して辞書の反復順で上書きされてしまう不具合の両方を修正して
    確立した。

    `tier_info`(id(row) -> (段ラベル, y_max))・`lengths`(id(row) -> 長さ)は
    呼び出し側で計算済みのものを渡すこと。戻り値は id(row) -> (left, right)。
    """
    raw_groups: dict[tuple, list] = {}
    for r in rows:
        key = (round(r.x_mark, 1), round(r.x_next, 1))
        raw_groups.setdefault(key, []).append(r)

    # 同じ列位置(x_mark, x_next)でも、繰り返しテンプレートにより全く別の
    # 段(離れたY位置)の製品が同じ列位置を再利用していることがある
    # (2026-09-14、WA2-1B-10のEA22-1TB489-1[1段]とEA22-1TB441-7[2段]、
    # WA2-1B-01のWA21-1B34-16[1段]とWA21-1B34-26[2段]で確認。列だけで
    # グループ化すると、本来別々の製品なのに継手の結果が強制的に同じに
    # なってしまっていた)。`find_product_rows_shared_group`の表検出と
    # 同様に、Yが大きく離れている(目安3000mm以上)メンバー同士は別の
    # グループとして分割する。
    groups: dict[tuple, list] = {}
    for key, members in raw_groups.items():
        members_by_y = sorted(members, key=lambda r: r.y)
        clusters: list[list] = [[members_by_y[0]]]
        for r in members_by_y[1:]:
            if r.y - clusters[-1][-1].y > 3000.0:
                clusters.append([r])
            else:
                clusters[-1].append(r)
        for i, cluster in enumerate(clusters):
            groups[(key, i)] = cluster

    centers = ex.find_stadium_centers(tdf)
    candidates_raw = []
    for mx, my, _r in centers:
        near = [rec for rec in tdf.texts if abs(rec.x - mx) < 3 and abs(rec.y - my) < 3]
        for rec in near:
            t = tdf.resolve_text(rec)
            if t:
                candidates_raw.append((mx, my, t))

    # (候補位置, テキスト, 側) ごとに、最も近いグループ(と距離)を記録する
    claims: dict[tuple, tuple] = {}
    for key, members in groups.items():
        member_lengths = [lengths[id(r)] for r in members if lengths.get(id(r)) is not None]
        if not member_lengths:
            continue
        length_value = max(member_lengths)
        main_axis_deg = 0.0
        theta = -math.radians(main_axis_deg)

        def rotate(x, y, theta=theta):
            xr = x * math.cos(theta) - y * math.sin(theta)
            yr = x * math.sin(theta) + y * math.cos(theta)
            return xr, yr

        # 長さに比例した閾値だけでは、極端に短い製品(300〜1500mm程度)で
        # 閾値が小さくなりすぎて、実在する継手候補(自分の設計符号と同じ
        # テキスト)まで届かないことがある。継手マークは製品の物理的な
        # 位置からほぼ一定の距離(実測で最大2908.5程度)に描かれるため、
        # 最低保証の閾値を設ける。グループ単位の相互排他により、閾値を
        # 広げても無関係な候補(別の設計符号のテキスト)を誤って奪う
        # リスクは無い(2026-09-14、EA1-1B-09のEA12-1B34-51・
        # EA12-1B29-52/53で確認)。
        threshold = max(length_value * 0.6, JOINT_MIN_THRESHOLD)
        for r in members:
            _tier_label, y_max = tier_info.get(id(r), (None, None))
            row_length = lengths.get(id(r))
            line_endpoints = None
            if row_length is not None:
                line_endpoints = _find_reference_line_endpoints(
                    tdf, r.x_mark, r.x_next, r.y, y_max, row_length,
                )
            if line_endpoints is not None:
                # 実際の部材線の左右端点を基準に、2次元距離(X・Yとも)で
                # 判定する。これにより、X方向だけ近いが実際にはY方向に
                # 大きく離れた(=別の部材の)候補を誤って採用しなくなる。
                (lx, ly), (rx, ry) = line_endpoints
                left_ref, right_ref = (lx, ly), (rx, ry)
                use_2d = True
            else:
                # 部材線が見つからない場合は既存方式(マーク/本数セル自身の
                # X座標、Yは考慮しない)にフォールバックする。
                x_mark_r, _ = rotate(r.x_mark, r.y)
                x_next_r, _ = rotate(r.x_next, r.y)
                left_ref, right_ref = (x_mark_r, None), (x_next_r, None)
                use_2d = False
            for mx, my, t in candidates_raw:
                if my <= r.y or (y_max is not None and my >= y_max):
                    continue
                xr, yr = rotate(mx, my)
                for side, ref in (("left", left_ref), ("right", right_ref)):
                    if use_2d:
                        dist = math.hypot(mx - ref[0], my - ref[1])
                    else:
                        dist = abs(xr - ref[0])
                    if dist > threshold:
                        continue
                    cand_key = (round(mx, 2), round(my, 2), t, side)
                    cur = claims.get(cand_key)
                    if cur is None or dist < cur[0]:
                        claims[cand_key] = (dist, key)

    # 1つのグループが複数の候補で「勝利」することがある(グループ側で他に
    # 競合が無ければ、距離に関係なく全部勝ってしまうため)。各グループ・
    # 各サイドにつき、実際に採用するのは距離が最小の1つだけにする。
    best_per_group_side: dict[tuple, tuple] = {}
    for (_mx, _my, t, side), (dist, key) in claims.items():
        cur = best_per_group_side.get((key, side))
        if cur is None or dist < cur[0]:
            best_per_group_side[(key, side)] = (dist, t)

    group_result: dict[tuple, list] = {}
    for (key, side), (_dist, t) in best_per_group_side.items():
        left_right = group_result.setdefault(key, [None, None])
        left_right[0 if side == "left" else 1] = t

    row_result = {}
    for key, members in groups.items():
        left, right = group_result.get(key, (None, None))
        for r in members:
            row_result[id(r)] = (left, right)
    return row_result
