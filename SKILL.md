---
name: tdf-master-extract
description: 実寸法師(.tdf)の鉄骨図面ファイルから製品マスタ情報(図番・製品マーク・設計符号・サイズ・本数・重量・長さ・左継手・右継手・種別・製品段)を抽出し、Excelにまとめるスキル。「実寸法師から抽出して」「tdfファイルからマスタを作って」「実寸法師の図面をExcel化して」「製品情報を一括で拾って」といった依頼、またはフォルダ内に.tdfファイルがあり実寸法師のデータをExcelにまとめたいという話題が出たときに使う。大梁(1G系、1マーク=1製品)・小梁(1B系、1製品に複数マークが並ぶ共有テーブル)の両パターンに対応し、普通梁(水平配置)・斜め梁(回転配置)のどちらにも対応する。DXFファイルからの抽出はこのスキルの対象外(dxf-masterスキルを使う)。
---

# 実寸法師(tdf)製品マスタ抽出スキル

実寸法師の.tdfファイル(バイナリ形式、公式な読み取りAPIは無い)を直接パースし、
図面上の製品情報テーブルと寸法・継手記号から、以下の項目を機械的に抽出する。

| 項目 | 抽出元 |
|---|---|
| 図番 | タイトル欄のテキスト |
| 製品マーク | 製品情報テーブル1列目 |
| 設計符号 | 製品情報テーブル2列目(小梁は複数マークで1つを共有することがある) |
| サイズ | 製品情報テーブル3列目(同上) |
| 本数 | 製品情報テーブル4列目(「N台」) |
| 重量 | 製品情報テーブル5列目(あれば。「N.Nkg」) |
| 長さ | 寸法値テキストと図形(直線)の突き合わせで推定 |
| 左継手・右継手 | 「長丸」記号で囲まれたテキストの位置関係から推定 |
| 種別 | 普通梁/斜め梁(部材の傾きから判定、小梁のみ出力) |
| 製品段 | 1枚の図面に立面が複数段描かれている場合の段数(小梁のみ出力) |

大梁(1G系: 1マーク=1製品)と小梁(1B系: 1製品に複数マークが並び設計符号・
サイズ・継手を共有する)は図面パターンが大きく異なるため、**完全に別の
スクリプト・別のロジックで実装されている**(大梁側は一切変更せず、小梁側を
新規追加する形で開発した)。

- 大梁: `scripts/tdf_master_extractor.py` + `scripts/extract_to_excel.py`
- 小梁: `scripts/tdf_master_extractor_multi.py`(大梁側の関数を再利用しつつ
  拡張) + `scripts/extract_to_excel_small_beam.py`

抽出ロジックは2026-09-11の普通梁・斜め梁2ファイルの検証から始まり、
2026-09-14に大梁48ファイル・小梁47+9ファイルの実データ突き合わせ検証を
経て確立したもの(すべてユーザー提供の正解データ・画像との照合による)。
バイナリのレコード構造・各判定ルールの根拠は`references/binary_format.md`
(主に大梁側)にまとめてある。挙動を変更・拡張する場合は必ず先にこれを
読むこと(特に「長さ」「継手」「段」の判定は一見単純な条件に見えて、
取付ピッチとの誤認識や、繰り返しテンプレートによる列位置の偶然の一致を
避けるための複数の除外・復活・分割ルールが積み重なっているため、根拠を
理解せずに簡略化すると精度が落ちる)。

## 使い方

### 大梁(1G系)のフォルダ一括抽出 → Excel出力

```
python scripts/extract_to_excel.py <tdfフォルダ> [<tdfフォルダ2> ...] <出力xlsxパス>
```

指定フォルダを再帰的に検索して全`*.tdf`を処理し、1シートの表(ファイル名列
付き)にまとめる。製品情報テーブルが見つからなかったファイルや例外が起きた
ファイルは、結果を止めずに「エラー」シートに記録する。

### 小梁(1B系)のフォルダ一括抽出 → Excel出力

```
python scripts/extract_to_excel_small_beam.py <tdfフォルダ> [<tdfフォルダ2> ...] <出力xlsxパス>
```

大梁版と異なり、**指定フォルダの直下の`*.tdf`のみを対象にする(再帰しない)**。
小梁詳細図フォルダには「00.旧」「01.発行歴」等の下層フォルダに古い版が
残っていることが多く、これらを取り違えて集計しないための仕様。

### 1ファイルだけ確認したい場合(大梁のみ)

```
python scripts/tdf_master_extractor.py <tdfパス>
```

標準出力に抽出結果を表示するだけの簡易確認用(小梁側には同等のCLIは無い)。

### コードから使う場合(大梁)

```python
import sys
sys.path.insert(0, "scripts")
import tdf_master_extractor as ex

tdf = ex.tb.load(path)
drawing_number = ex.find_drawing_number(tdf)
rows = ex.find_product_rows(tdf)          # 製品マーク・設計符号・サイズ・本数・重量
for row in rows:
    length, debug = ex.determine_length(tdf, row.x_mark, row.x_next, row.y)
    left, right = ex.determine_joints(
        tdf, row.x_mark, row.x_next, row.y, debug["main_axis_deg"], length
    )
```

### コードから使う場合(小梁)

```python
import sys
sys.path.insert(0, "scripts")
import tdf_master_extractor as ex
import tdf_master_extractor_multi as exm
import extract_to_excel_small_beam as exb

tdf = ex.tb.load(path)
rows = exb.get_rows(tdf)          # 1マーク=1製品 + 共有テーブルの両パターンをマージ
exb.sort_rows(rows)
tier_info = exb.compute_tier_info(rows)   # 段(1段/2段…)とY境界を計算

lengths = {}
for row in rows:
    _tier_label, y_max = tier_info.get(id(row), (None, None))
    length_value, debug = ex.determine_length(tdf, row.x_mark, row.x_next, row.y, y_max=y_max)
    lengths[id(row)] = length_value

joints = exm.assign_joints_batch(tdf, rows, tier_info, lengths)
```

## 実行前の確認事項

- 対象フォルダに`.tdf`ファイルが実際に存在するか、Windows側で
  `Get-ChildItem -Recurse -Filter *.tdf`等を使って事前に確認する
  (UNCパス・日本語パスが多いので、PowerShellツールから実行すること。
  Bash/Git Bashはバックスラッシュを解釈してUNCパスを壊すため避ける)。
- 出力Excelのシート名・列名は固定(`抽出結果`シート)。大梁は
  ファイル名/図番/製品マーク/設計符号/サイズ/本数/長さ/重量/左継手/右継手、
  小梁はさらに種別/製品段が加わる。列の追加・並び替えの要望があれば
  `extract_to_excel.py`/`extract_to_excel_small_beam.py`の`HEADERS`と
  `ws.append(...)`の対応を書き換える。
- 対象の図面が大梁(1G系)か小梁(1B系)かを取り違えないこと。両者は
  マーク名の命名規則(`1G-`系か`1B-`系か)や実際のフォルダ名(「大梁詳細図」
  「小梁詳細図」等)で判別できる。

## 結果の伝え方

- 抽出できた行数、エラーになったファイル数を必ず報告する。
- 「長さ」「左継手」「右継手」がNone(空欄)になった製品は、根拠となる
  寸法テキストや長丸記号が図面から見つからなかったことを意味する。
  0件だからといって黙って握りつぶさず、どのファイル・どの製品で
  抽出できなかったかをユーザーに伝える。
- この抽出ロジックは実データ(大梁48ファイル・小梁47+9ファイル)で検証
  済みだが、`references/binary_format.md`の「既知の限界」に記載の通り、
  未検証のパターン(製品テーブルの列構成違い・継手コードのプレフィックス
  違い等)は残っている。明らかに様子の違う図面で結果がおかしい場合は、
  機械的に「バグ」と決めつけず、まずユーザーに実際のCAD画面(または
  スクリーンショット)を見せてもらい、判定ルールの前提が崩れていないか
  確認する。**tdfファイルを実寸法師CADで実際に開いて確認する操作は
  行わない**(誤操作による上書き保存のリスクを避けるため)。バイナリを
  直接解析し、ユーザー提供のスクリーンショットや正解データで照合する
  方式を徹底すること。
