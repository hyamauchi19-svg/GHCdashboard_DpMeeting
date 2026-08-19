---
name: shinryoka-mdc6-csv-to-xlsx
description: shinryoka-betsu-jisseki スキル(病院ダッシュボードχ)でダウンロードしたMDC6別・症例一覧CSV(診療科ごと)を、過去年度の実績フォルダ(例: 2025実績)にある既存Excelファイルと同じフォーマット(1行目タイトル・列数・列幅・フォント・オートフィルター・症例数の降順ソート)に変換し、当年度の実績フォルダ(例: 2026実績)にxlsxとして保存する定型フロー。ユーザーが「CSVをExcelに変換して」「症例一覧をExcel化して」「◯◯実績フォルダのファイルのように整形して」のように、病院ダッシュボードχ由来のMDC6別CSVをフォーマット揃えでExcel化してほしいと頼んだときに使う。このPCにはPythonが入っていないため、xlsxスキル標準のopenpyxlではなく、インストール済みのMicrosoft Excelを`PowerShell`のCOM自動化で操作して作成する。
---

# 症例一覧CSV → Excel変換 (前年度フォーマット踏襲)

`shinryoka-betsu-jisseki` スキルでダウンロードした「MDC6別・年度」の詳細データCSV(診療科ごとに1ファイル)を、前年度の実績フォルダにある既存Excelファイルと**見た目・列構成を完全に揃えて**当年度の実績フォルダへExcel(.xlsx)として保存する。

## 前提: このPCにはPythonが無い

`xlsx` スキルの標準手順(`openpyxl`/`pandas`/`markitdown`)は、Python自体がこのマシンに存在しないため使えない(`which python` で見つかるのは Microsoft Store 誘導用のスタブで実行不可)。**必ず先にPythonの有無を確認し、無ければ本スキルのExcel COM自動化(PowerShell)に切り替える。** 逆にPythonが使える環境であれば、素直に `xlsx` スキルの `openpyxl` を使った方が速く確実なので、そちらを優先してよい。

## 前提2: Excel COMは「新規作成」はできるが「既存ファイルを開く」のは失敗することがある

このマシン(非対話的なセッションでPowerShellツールが動く環境)では、次の非対称な挙動が確認されている:

- `New-Object -ComObject Excel.Application` → `$excel.Workbooks.Add()` (真っ白な新規ブックの作成) は **成功する**。
- `$excel.Workbooks.Open($path)` (既存ファイルを開く) は、パスの文字種やファイルの中身に関係なく **一貫して失敗する**(`Workbooks クラスの Open プロパティを取得できません。` / HRESULT `0x800A03EC`)。Protected View・自動化セキュリティ・リンク更新確認などの引数を色々変えても変わらない。原因はおそらく非対話セッション(window station/デスクトップ相互作用なし)によるUI依存処理の失敗で、環境側の制約と考えられる。

したがって:
- **前年度ファイルなど既存xlsxの構造を確認したいとき、Excel COMで開こうとしない。** xlsxはzip形式なので、Git Bashの `unzip` コマンドで展開し、`xl/sharedStrings.xml`(文字列一覧)・`xl/worksheets/sheet1.xml`(セル配置・書式ID・オートフィルター・ソート状態)・`xl/styles.xml`(フォント・罫線・数値書式)を直接読む。
- **新規に作るExcelファイルはWorkbooks.Add()で真っ白なブックを作り、そこにCOMでセル・書式を書き込んでSaveAsする。**
- もし将来的にWorkbooks.Openが成功する環境(対話的セッション、別マシン等)であれば、素直に開いて読み書きしてよい。まず `Add()` と `Open()` の両方を1回ずつ試し、`Open` が通るならそちらの方が構造把握が楽なので活用する。

## 操作手順

1. **対象フォルダとファイルの特定**
   ユーザーが指す「前年度実績フォルダ」(例: `2025実績`)と「当年度実績フォルダ」(例: `2026実績`、多くの場合空フォルダとして既に用意されている)を確認する。`find`/`ls` で両フォルダのファイル一覧を突き合わせ、変換元CSV(通常 `shinryoka-betsu-jisseki` の実行結果、診療科ごとに1ファイル)と前年度ファイルが科ごとに1:1対応しているか確認する。ファイル数が合わない場合は科の過不足をユーザーに確認する。

2. **前年度ファイル1件の構造をunzipで確認する**
   ```bash
   mkdir -p /tmp_or_scratchpad/xlsx_inspect
   unzip -o "前年度ファイル.xlsx" -d /tmp_or_scratchpad/xlsx_inspect/ref
   cat xl/sharedStrings.xml        # 1行目タイトル文言・全文字列を確認
   cat xl/worksheets/sheet1.xml    # dimension(範囲)、行1のセル配置(s=スタイルID)、autoFilter、sortState を確認
   cat xl/styles.xml               # cellXfs の numFmtId(3=#,##0、10=0.00% など)、フォント名・サイズ、列幅(cols)、行高(row ht)
   ```
   ここで必ず確認すべき点:
   - **1行目(タイトル行)の実際の文言と列数** — 元CSVの列とそのまま1:1ではなく、削られている列(例: 1日単価・1日平均患者数)や、ラベルが意味的に読み替えられている列(例: 「年度」列の見出しが実際には「4～7月」のような期間ラベルになっている)があり得る。**必ず sharedStrings.xml の該当インデックスを row1 のセル参照(`t="s"><v>N</v>`)と突き合わせて確定させる。推測しない。**
   - **除外されている行があるか** — 元CSVには含まれる「NULL」(MDC6未分類)行が、前年度ファイルのsharedStringsに存在しなければ、除外対象と判断する。
   - **ソート順** — `<sortState ref="..."><sortCondition descending="1" ref="D1:D..."/></sortState>` のように、どの列で降順/昇順ソートされているか。
   - **数値書式** — `cellXfs` の `numFmtId` (3=`#,##0`、10=`0.00%` が典型)。どの列がどの書式か、`sheetData` 内の各セルの `s="N"` から対応付ける。
   - **列幅・行高・フォント・ウィンドウ枠固定** — `cols`(列幅)、`row r="1" ht="..."`、`fonts`(フォント名・サイズ)、`sheetView`内の`pane`(freeze panes)。

3. **PowerShellでExcel COMを使い、ブックを新規作成してデータを書き込む**
   1科ずつ以下を行う(複数科ある場合は同一Excelインスタンスを使い回してループするとオーバーヘッドが少ない):
   - `Import-Csv -LiteralPath $csvPath -Encoding UTF8 -Header @(...)` で列名を明示指定しつつ読み込む(元CSVの1行目はヘッダーなので `-Header` を指定すると1行目がデータ扱いされてしまう点に注意。`Select-Object -Skip 1` で読み飛ばす)。
   - 除外対象行(例: 病名列が `"NULL"`)を `Where-Object` でフィルタ。
   - 金額・件数列はカンマを取り除いて `[double]`/`[int]` にキャスト。パーセント文字列(`"33.3%"`)は `TrimEnd('%')` して100で割り小数(フラクション)に変換してから書き込む(空文字列ならセルに書き込まず空欄のままにする)。
   - `$wb = $excel.Workbooks.Add()` → `$ws = $wb.Sheets.Item(1)` → ヘッダー行・データ行をセルに直接代入。
   - 数値書式は `Range.NumberFormat = "#,##0"` / `"0.00%"` のように文字列で指定すればよい(手順2で確認したnumFmtIdに対応する書式文字列を使う)。
   - 列幅は `ws.Columns.Item(N).ColumnWidth = 数値`、行高は `ws.Rows.Item(1).RowHeight = 数値`、フォントは `Range.Font.Name` / `.Size` で設定。
   - **ソートはExcel自身の `Range.Sort()` メソッドで実行する**(手作業で事前ソート済みデータを書くのではなく、書き込み後にExcelでソートさせると、オートフィルターの「現在の並び替え状態」インジケーターも自動的に整合する)。例: `$sortRange.Sort($ws.Range("D1:D$lastRow"), 2, [Type]::Missing, [Type]::Missing, 1, [Type]::Missing, [Type]::Missing, 1)` (第2引数2=降順、Header=1=あり)。
   - オートフィルターは `$ws.Range("A1:I$lastRow").AutoFilter()`。
   - ウィンドウ枠固定(先頭行・先頭列)は `$ws.Range("B2").Select()` してから `$excel.ActiveWindow.FreezePanes = $true`。
   - `$wb.SaveAs($outPath, 51)` (51 = `xlOpenXMLWorkbook`、拡張子.xlsx) で保存し、`$wb.Close($false)`。
   - 全科処理後に `$excel.Quit()` と `[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null` を必ず実行し、Excelプロセスを残さない。

4. **出力ファイル名の決定**
   前年度ファイルの命名規則(例: `01：救急科症例一覧2025.xlsx` のように全角コロン区切り)をそのまま踏襲し、年度部分だけ当年度に置き換える。CSV側が `01_救急科症例一覧2026.csv` のようにアンダースコア区切りであれば、`-replace '^(\d+)_', '$1：'` のように先頭のアンダースコアだけ全角コロンに置換し、拡張子を`.xlsx`にする。

5. **生成後の検証**
   Excel COMで開いて確認するのではなく(手順2と同じ理由でOpenが失敗しうる)、**生成した.xlsxもunzipして中身を確認する**。最低限:
   - `dimension` が想定行数(ヘッダー+データ行数)と一致しているか。
   - `sharedStrings.xml` の先頭に1行目タイトルが意図通り入っているか。
   - データ行が症例数降順になっているか(`sheetData`の該当列の値を目視で降順確認)。
   - `NULL`行が含まれていないか。
   1件検証してから残りをまとめて処理するとやり直しが少ない。

6. **完了報告**
   何件変換し、どこに保存したか、除外した行(NULL行など)や列数調整の内容をユーザーに簡潔に報告する。

## つまずきやすい点

- **`New-Object -ComObject Excel.Application` を複数回失敗させると、タイトルなしの非表示Excelプロセスが残ることがある。** `Get-Process -Name EXCEL | Select Id,StartTime,MainWindowTitle` で確認し、`MainWindowTitle` が空でユーザーの作業中ドキュメントでないことを確認できれば `Stop-Process -Force` で片付けてから再試行する(ユーザーが実際にExcelを開いて作業中の可能性がある場合は先に確認する)。
- CSVの数値列はカンマ区切り文字列("1,540")なので、そのまま`[double]`にキャストするとエラーになる。必ず `-replace ',',''` してからキャストする。
- パーセント列が空文字列("")の行がある(平均在院日数が短く「期間Ⅱ超率」が定義されない場合など)。空のまま書き込もうとして`[double]""`のようなキャストをするとエラーになるので、事前に空文字チェックしてスキップする。
- `Import-Csv -Header` を使うと元CSVの1行目(本来のヘッダー)が最初のデータレコードとして読み込まれてしまう。`Select-Object -Skip 1` を忘れない。
- PowerShellの `Where-Object`/`Select-Object` の結果が1件だけの場合、配列ではなく単一オブジェクトになり `.Count` が `$null` になる(件数ログ表示が空になるだけで処理自体は正常)。1件だけの科がある場合は驚かないこと。`foreach`ループはコレクションが1件でも正しく1回実行される。
- 前年度ファイルの1行目タイトルは、元データの列名をそのまま転記したものとは限らない(例: 「年度」列の見出しが実際には対象期間を表す「4～7月」等になっている)。**必ず手順2でsharedStrings.xmlを見て実際の文言を確認し、それを使う。前回確認した値を次回もそのまま使い回してよいとは限らない(年ごとに書き換えているファイルもあり得る)。**
