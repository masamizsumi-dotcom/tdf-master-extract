"""小梁(1B系)専用の抽出→Excel出力スクリプト。

大梁(1G系、`extract_to_excel.py`)とは以下の点で処理が異なるため、
既存コードには一切手を加えず、こちらを完全に別スクリプトとして用意した
(2026-09-14、ユーザー指示: 「1マーク=1製品」の大梁パターンは今まで通り、
「1製品に複数マークがある」小梁パターンは新しいコードで)。

  - 製品マーク検出: `tdf_master_extractor.find_product_rows`(1マーク=1製品)
    に加えて`tdf_master_extractor_multi.find_product_rows_shared_group`
    (複数マークが設計符号・サイズを共有するパターン)を併用する。
  - 並び順: 上にある製品が先→同じ高さなら左の製品が先→同じ製品内は
    マーク末尾の番号昇順。
  - 長さ・継手の判定範囲: 1枚の図面に梁の立面が2段に描かれている場合、
    下段は「自分の段から直上の段まで」にY範囲を打ち切る(製品段列に
    1段/2段を記録する)。
  - 継手判定: 大梁の`GJ`のような専用プレフィックスが無いため、長丸に
    囲まれたテキストは全て継手マークの一種とみなし、位置(距離・段の
    Y範囲)だけで、かつグループ単位の相互排他(1つの継手候補は本当に
    最も近いグループにしか渡さない)で判定する
    (`tdf_master_extractor_multi.assign_joints_batch`)。

使い方:
    python extract_to_excel_small_beam.py <tdfフォルダ> [<tdfフォルダ2> ...] <出力xlsxパス>

大梁版と異なり、指定フォルダの直下の*.tdfのみを対象にする(再帰しない)。
小梁詳細図フォルダには「00.旧」「01.発行歴」等の下層フォルダに古い版が
残っていることが多く、これらを取り違えて集計しないようにするため。
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import tdf_master_extractor as ex
import tdf_master_extractor_multi as exm
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

HEADERS = [
    "ファイル名", "図番", "製品マーク", "設計符号", "サイズ", "本数",
    "長さ(m)", "重量", "左継手", "右継手", "種別", "製品段",
]

_AXIS_TOLERANCE_DEG = 2.0

# 1枚の図面に梁の立面が2段(以上)に描かれている場合、下段の行から見て
# 上段の行より上は別の梁の寸法線・継手なので含めてはいけない。
# 同一グループ内の行同士のY間隔(実測240前後)と、段違いのY間隔(実測7700
# 以上)の間に十分な余裕を持たせた係数として3000を採用する
# (2026-09-14、EA1-1B-01のB1000A/B90A2グループで確認)。
TIER_Y_GAP_THRESHOLD = 3000.0


def classify_beam_type(main_axis_deg: float | None) -> str | None:
    if main_axis_deg is None:
        return None
    remainder = abs(main_axis_deg) % 90.0
    if remainder <= _AXIS_TOLERANCE_DEG or remainder >= 90.0 - _AXIS_TOLERANCE_DEG:
        return "普通梁"
    return "斜め梁"


def get_rows(tdf) -> list:
    rows = ex.find_product_rows(tdf)
    existing = {(round(r.x_next, 1), round(r.y, 1)) for r in rows}
    rows = rows + exm.find_product_rows_shared_group(tdf, existing_positions=existing)
    # ex.find_product_rows(1マーク=1製品パターン)には削除マーク(×印)の
    # チェックが元々無いため、行の出どころに関わらずここでまとめて除外する。
    rows = exm.filter_deleted_rows(tdf, rows)
    return rows


def sort_rows(rows: list) -> None:
    """上にある製品が先→同じ高さなら左の製品が先→同じ製品(マーク接頭辞が
    同じもの)内はマーク番号の昇順(2026-09-14指示)。「同じ製品」は
    マーク末尾の番号を除いた接頭辞が一致するマーク同士とみなす(物理的に
    複数の表・列に分かれて描かれていても、同じ接頭辞ならまとめて番号順に
    並べる。EA1-1B-06でB34マークが複数の表に分かれていた事例で確認)。"""
    prefix_max_y: dict[str, float] = {}
    prefix_min_x: dict[str, float] = {}
    for r in rows:
        p = exm.mark_prefix(r.mark)
        prefix_max_y[p] = max(prefix_max_y.get(p, r.y), r.y)
        prefix_min_x[p] = min(prefix_min_x.get(p, r.x_mark), r.x_mark)

    def key_func(r):
        p = exm.mark_prefix(r.mark)
        return (-prefix_max_y[p], prefix_min_x[p], exm.mark_sort_key(r.mark))

    rows.sort(key=key_func)


def compute_tier_info(rows: list) -> dict:
    """列(X座標)に関係なく、ファイル全体を対象にY座標で「段」クラスタリング
    する。製品マーク群のY座標が(TIER_Y_GAP_THRESHOLD以上)大きく異なる
    グループを、ページの上半分=1段・下半分=2段とみなす(2026-09-14、
    ユーザー確認: 段はX座標の列ではなくページ全体のY位置で決まる。
    EA1-1B-09で、列が異なるEA12-1B34-13[上]とEA12-1B29-51/52/53[下]が
    同じ2段になる実例で確認)。
    戻り値は id(row) -> (段ラベル, 長さ・継手判定に使うy_max)。"""
    if not rows:
        return {}

    sorted_rows = sorted(rows, key=lambda r: -r.y)
    tiers = [[sorted_rows[0]]]
    for r in sorted_rows[1:]:
        if tiers[-1][-1].y - r.y > TIER_Y_GAP_THRESHOLD:
            tiers.append([r])
        else:
            tiers[-1].append(r)

    info = {}
    for i, tier in enumerate(tiers):
        tier_label = f"{i + 1}段"
        if i == 0:
            for r in tier:
                info[id(r)] = (tier_label, None)
            continue

        prev_tier = tiers[i - 1]
        global_y_max = min(r.y for r in prev_tier)
        for r in tier:
            # y_maxは本来「直上の段で自分と同じ列にある行」までで区切る
            # べきだが、直上の段全体(列を問わない)の最小Yを使っていたため、
            # 無関係な別列がたまたま直上の段で最も小さいYを持つ場合に、
            # 本来の(同じ列の)寸法線・継手より手前で範囲が打ち切られて
            # しまう不具合があった(2026-09-14、EA2-1B-15のEA22-1TB489-9で、
            # 直上の段のEA22-1TB489-1[別列]が最小Yだったため範囲が狭くなり、
            # 本来同じ列で共有しているEA22-1TB441側の寸法線[3205]が範囲外
            # になって長さ判定不能になっていた。ユーザーの指摘により、
            # X範囲が重なる[=同じ列とみなせる]行だけを対象にY境界を計算する
            # よう修正。該当する行が無い場合は従来通り直上の段全体の最小Y
            # にフォールバックする)。
            same_col = [
                pr for pr in prev_tier
                if max(r.x_mark, pr.x_mark) < min(r.x_next, pr.x_next)
            ]
            y_max = min(pr.y for pr in same_col) if same_col else global_y_max
            info[id(r)] = (tier_label, y_max)
    return info


_FILENAME_NUM_PATTERN = re.compile(r"^(.*-)(\d+)([A-Za-z]*)(.*)$")


def _filename_sort_key(path: Path):
    """ファイル名を自然順(数値として)ソートするためのキー。

    単純な文字列比較では、`WA1-1B-02A`(区切り文字が半角/全角スペースより
    小さい`A`)が`WA1-1B-02`より前に来てしまう(2026-09-14、ユーザー
    指摘: 図番の並びは-02の次に-02Aであるべき)。図番末尾の「-数字+
    英字接尾辞」を切り出し、(prefix, 数値, 接尾辞, 残り)で比較することで
    -02の後に-02Aが続く自然な順序にする。パターンに一致しない場合は
    従来通りファイル名の文字列比較にフォールバックする。"""
    name = path.name
    m = _FILENAME_NUM_PATTERN.match(name)
    if not m:
        return (name, 0, "", "")
    prefix, num, suffix, rest = m.groups()
    return (prefix, int(num), suffix, rest)


def collect_tdf_files(folder: Path) -> list[Path]:
    return sorted((p for p in folder.glob("*.tdf") if p.is_file()), key=_filename_sort_key)


def main() -> None:
    if len(sys.argv) < 3:
        print("使い方: python extract_to_excel_small_beam.py <tdfフォルダ> [<tdfフォルダ2> ...] <出力xlsxパス>")
        raise SystemExit(1)

    folders = [Path(p) for p in sys.argv[1:-1]]
    out_path = Path(sys.argv[-1])

    tdf_files: list[Path] = []
    for folder in folders:
        if not folder.exists():
            print(f"フォルダが見つかりません: {folder}")
            raise SystemExit(1)
        tdf_files.extend(collect_tdf_files(folder))
    print(f"対象ファイル数: {len(tdf_files)}")

    wb = Workbook()
    ws = wb.active
    ws.title = "抽出結果"
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")

    errors = []
    total_rows = 0
    for tdf_path in tdf_files:
        try:
            tdf = ex.tb.load(str(tdf_path))
            drawing_number = ex.find_drawing_number(tdf)
            rows = get_rows(tdf)
            if not rows:
                errors.append((tdf_path.name, "製品情報テーブルが見つかりませんでした"))
                continue

            sort_rows(rows)
            tier_info = compute_tier_info(rows)

            lengths: dict[int, float | None] = {}
            beam_types: dict[int, str | None] = {}
            for row in rows:
                _tier_label, tier_y_max = tier_info.get(id(row), (None, None))
                if len(rows) == 1:
                    x_min = row.x_mark - ex.RELAX_MARGIN
                    x_max = row.x_next + ex.RELAX_MARGIN
                else:
                    x_min = row.x_mark
                    x_max = row.x_next
                length_value, debug = ex.determine_length(tdf, x_min, x_max, row.y, y_max=tier_y_max)
                lengths[id(row)] = length_value
                beam_types[id(row)] = classify_beam_type(
                    debug.get("main_axis_deg") if length_value is not None else None
                )

            joints = exm.assign_joints_batch(tdf, rows, tier_info, lengths)

            for row in rows:
                tier_label, _y_max = tier_info.get(id(row), (None, None))
                length_value = lengths[id(row)]
                left, right = joints.get(id(row), (None, None))
                if length_value is None:
                    left = right = None
                # 記録はm単位(内部の判定ロジックはmm前提のまま、出力直前だけ変換)。
                length_value_m = None if length_value is None else length_value / 1000.0
                ws.append([
                    tdf_path.name, drawing_number, row.mark, row.design_code,
                    row.size, row.count, length_value_m, row.weight, left, right,
                    beam_types[id(row)], tier_label,
                ])
                total_rows += 1
        except Exception as e:  # noqa: BLE001 一括処理のため個別ファイルの失敗で全体を止めない
            errors.append((tdf_path.name, f"{type(e).__name__}: {e}"))

    for col_idx, header in enumerate(HEADERS, start=1):
        width = max(12, len(header) * 2)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    if errors:
        ws_err = wb.create_sheet("エラー")
        ws_err.append(["ファイル名", "内容"])
        for cell in ws_err[1]:
            cell.font = Font(bold=True)
        for name, msg in errors:
            ws_err.append([name, msg])
        ws_err.column_dimensions["A"].width = 40
        ws_err.column_dimensions["B"].width = 60

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    print(f"完了: {total_rows}行を抽出しました({len(errors)}件のエラー)")
    print(f"出力先: {out_path}")


if __name__ == "__main__":
    main()
