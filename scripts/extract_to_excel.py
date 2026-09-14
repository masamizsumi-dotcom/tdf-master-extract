"""指定フォルダ内の.tdfファイルをすべて解析し、抽出結果を1つのExcelにまとめる。

使い方:
    python extract_to_excel.py <tdfフォルダ> [<tdfフォルダ2> ...] <出力xlsxパス>

フォルダ内の*.tdfを再帰的に検索し、ファイルごとに製品マスタ情報
(図番・製品マーク・設計符号・サイズ・本数・重量・長さ・左継手・右継手・
種別(普通梁/斜め梁))を抽出して1シートの表にまとめる。複数フォルダを
指定した場合は指定順にまとめて1つの表に出力する。抽出できなかった項目は
空欄で出力し、処理自体が失敗したファイルはエラー内容を別シート「エラー」
に記録する。
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
    "長さ", "重量", "左継手", "右継手", "種別", "製品段",
]

_AXIS_TOLERANCE_DEG = 2.0

# 1枚の図面に梁の立面が2段(以上)に描かれている場合、下段の行から見て
# 上段の行より上は別の梁の寸法線なので長さ判定に含めてはいけない。
# 同一グループ内の行同士のY間隔(実測240前後)と、段違いのY間隔(実測7700
# 以上)の間に十分な余裕を持たせた係数として3000を採用する
# (2026-09-14、EA1-1B-01のB1000A/B90A2グループで確認)。
TIER_Y_GAP_THRESHOLD = 3000.0



def sort_rows(rows: list) -> None:
    """製品の並び順を「上にある製品が先→同じ高さなら左の製品が先→同じ
    製品内はマーク番号の昇順」にする(2026-09-14、EA1-1B-01でB90A[上]→
    B1000A[下]の順になるよう指示を受けて追加)。

    「同じ製品」はマーク末尾の番号を除いた接頭辞が一致するマーク同士と
    みなす(物理的に複数の表・列に分かれて描かれていても、同じ接頭辞なら
    まとめて番号順に並べる。2026-09-14、EA1-1B-06でB34マークが複数の表に
    分かれていた事例で確認)。その製品の代表Y座標は最大Y(最も上にある
    行のY)を使う。"""
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

    戻り値は id(row) -> (段ラベル, 長さ判定に使うy_max) の辞書。
    Yが最も高い(図面上で最も上にある)段を「1段」とし、以降下に行くほど
    「2段」「3段」…と番号を振る。最上段(1段)はy_max=None(従来通り上方向
    無制限)。それ以外の段は、直上の段の最小Y座標をy_maxとして長さ判定の
    範囲を打ち切る(自分の段の寸法線が、直上の段の寸法線と混ざらない
    ようにするため)。"""
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
        y_max = None if i == 0 else min(r.y for r in tiers[i - 1])
        for r in tier:
            info[id(r)] = (tier_label, y_max)
    return info


def classify_beam_type(main_axis_deg: float | None) -> str | None:
    """主軸方向(main_axis_deg)から「普通梁」(0度/90度=図面の縦横に沿う
    向き)か「斜め梁」(それ以外の角度)かを判定する。長さが判定できず
    main_axis_degが無い場合はNoneを返す。"""
    if main_axis_deg is None:
        return None
    remainder = abs(main_axis_deg) % 90.0
    if remainder <= _AXIS_TOLERANCE_DEG or remainder >= 90.0 - _AXIS_TOLERANCE_DEG:
        return "普通梁"
    return "斜め梁"

_SEQ_NUMBER_PATTERN = re.compile(r"-(\d+)([A-Za-z]*)　")


def _natural_sort_key(name: str) -> tuple:
    """ファイル名の連番部分(例: "-05"や"-05A")を数値として比較し、
    英字無し("-05")が英字付き("-05A")より先に来るようにする自然順ソート
    キー(2026-09-14、WA2-1G-05/WA2-1G-05Aの並び順修正で追加)。"""
    m = _SEQ_NUMBER_PATTERN.search(name)
    if not m:
        return (name,)
    prefix = name[:m.start()]
    number = int(m.group(1))
    suffix = m.group(2)
    rest = name[m.end():]
    return (prefix, number, suffix, rest)


def collect_tdf_files(folder: Path) -> list[Path]:
    return sorted(folder.rglob("*.tdf"), key=lambda p: _natural_sort_key(p.name))


def main() -> None:
    if len(sys.argv) < 3:
        print("使い方: python extract_to_excel.py <tdfフォルダ> [<tdfフォルダ2> ...] <出力xlsxパス>")
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
            rows = ex.find_product_rows(tdf)
            # 「1製品に複数マークが共有設計符号・サイズで並ぶ」パターン(小梁等)は
            # 別コード(tdf_master_extractor_multi)で検出し、既存ロジックが拾った
            # daiセルとは重複しないようにマージする(2026-09-14追加)。
            existing_dai_positions = {(round(r.x_next, 1), round(r.y, 1)) for r in rows}
            rows = rows + exm.find_product_rows_shared_group(tdf, existing_positions=existing_dai_positions)
            sort_rows(rows)
            tier_info = compute_tier_info(rows)
            for row in rows:
                tier_label, tier_y_max = tier_info.get(id(row), (None, None))
                if len(rows) == 1:
                    x_min = row.x_mark - ex.RELAX_MARGIN
                    x_max = row.x_next + ex.RELAX_MARGIN
                else:
                    x_min = row.x_mark
                    x_max = row.x_next
                length_value, debug = ex.determine_length(tdf, x_min, x_max, row.y, y_max=tier_y_max)
                left = right = None
                if length_value is not None:
                    main_axis_deg = debug.get("main_axis_deg", 0.0)
                    # 継手判定は緩和前の製品自身のマーク/本数セル座標を基準にする
                    left, right = ex.determine_joints(
                        tdf, row.x_mark, row.x_next, row.y, main_axis_deg, length_value,
                        y_max=tier_y_max,
                    )
                beam_type = classify_beam_type(debug.get("main_axis_deg"))
                ws.append([
                    tdf_path.name, drawing_number, row.mark, row.design_code,
                    row.size, row.count, length_value, row.weight, left, right,
                    beam_type, tier_label,
                ])
                total_rows += 1
            if not rows:
                errors.append((tdf_path.name, "製品情報テーブルが見つかりませんでした"))
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
