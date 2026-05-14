---
name: invoice-teams
description: >
  注文書Excel（.xlsx）から請求書Excel（.xlsx）を自動生成するスキル。
  ユーザーが注文書（Excel）と請求書テンプレート（Excel）をアップロードして
  「請求書を作成して」「注文書から請求書に転記して」「自動入力して」などと言ったときに必ず使用すること。
  注文書の宛名会社・担当者・件名・明細（内容・数量・単価）を読み取り、
  請求書テンプレートの正しいセルに自動入力して完成した請求書ファイルを出力する。
---

# 注文書Excel → 請求書Excel 自動生成スキル

## 前提知識：注文書と請求書の転記ルール

注文書（JR東日本企画フォーマット）の各項目は以下のルールで請求書テンプレートに転記する。

| 注文書の項目 | 請求書テンプレートのセル | 備考 |
|---|---|---|
| 発注元会社名 | A5 | 「セールス番号／経費番号」行の列44(AU列)から取得 |
| 担当者名 | A6 | 「〇〇担当者」ラベルを含む行の列44から取得 ＋「様」を付与 |
| 注文件名 | A9 | 「注文件名」行の列14(O列)から取得 |
| 発行日 | G3 | スクリプト実行当日の日付 |
| 請求書番号 | G4 | **空欄のまま**（手入力） |
| 明細 内容 | A17〜A45 | 注文書のNO列が数字かつ内容がある行を抽出。文頭に「■」を付与 |
| 明細 数量 | D17〜D45 | 「1式」のような表記から数値のみ抽出（例：1式→1） |
| 明細 単価 | F17〜F45 | 注文書の単価をそのまま転記 |
| 金額・消費税・合計 | G列（数式セル） | テンプレートのExcel数式が自動計算するため入力不要 |

### 宛名会社（請求先）の抽出ルール（重要）
請求書の宛名会社は「TEAMSが請求を送る先＝発注元会社」である。
注文書の「セールス番号／経費番号」行の列44（AU列）に発注元会社名が入っている。
例：「株式会社ジェイアール東日本企画」

**注意**: 注文書の3行目には「* 株式会社　TEAMS（701227）御中」とあるが、
これは注文書の宛先（TEAMSへの発注通知）であり、請求書の宛名ではない。
請求書の宛名は必ず「セールス番号」行から取得すること。

### 担当者名の抽出ルール（重要）
注文書によって発注元会社名・担当者名のラベルが変わる。例：
- 「ジェイアール東日本企画担当者」
- 「〇〇会社担当者」など

そのため、**「担当者」という文字列を含むラベル**を動的に検索して取得すること。
担当者名も同行の列44（AU列）から取得する。

## 手順

### Step 1: ファイルの確認

ユーザーが以下をアップロードしているか確認する：
1. **注文書 Excel**（`.xlsx`）— 読み取り対象
2. **請求書テンプレート Excel**（`.xlsx`）— 入力対象

どちらかが足りない場合はアップロードを依頼する。

### Step 2: 注文書Excelの読み取り

`scripts/extract_po_excel.py` を実行して注文書から情報を抽出する：

```bash
python /path/to/skill/scripts/extract_po_excel.py <注文書Excelパス>
```

スクリプトがない場合は以下のPythonコードで抽出する：

```python
import openpyxl

def extract_from_excel(po_path):
    wb = openpyxl.load_workbook(po_path, data_only=True)
    # 「注文書」シートを優先、なければ最初のシートを使用
    ws = wb['注文書'] if '注文書' in wb.sheetnames else wb.active

    data = {
        '宛名会社': '',
        '担当者': '',
        '件名': '',
        '明細': [],
    }

    for row in ws.iter_rows(min_row=1, max_row=50, values_only=True):
        non_none = {i: v for i, v in enumerate(row) if v is not None}
        if not non_none:
            continue

        col0 = str(non_none.get(0, ''))

        # 発注元会社名：「セールス番号」行の列44
        if 'セールス番号' in col0:
            data['宛名会社'] = str(non_none.get(44, '')).strip()

        # 担当者名：「担当者」を含むラベルがある行の列44
        # ラベルは注文書によって異なる（例：「ジェイアール東日本企画担当者」「〇〇担当者」）
        for col_idx, val in non_none.items():
            if isinstance(val, str) and '担当者' in val and col_idx != 44:
                # 担当者ラベルの右側の列44に担当者名がある
                if 44 in non_none:
                    data['担当者'] = str(non_none[44]).replace('\u3000', ' ').strip()

        # 注文件名：「注文件名」行の列14
        if col0 == '注文件名':
            data['件名'] = str(non_none.get(14, '')).strip()

        # 明細：NO列が数字文字列かつ内容（列2）がある行
        no_val = non_none.get(0)
        if isinstance(no_val, (str, int)):
            no_str = str(no_val).strip()
            if no_str.isdigit() and non_none.get(2):
                qty_raw = str(non_none.get(41, '1式'))
                qty_num = ''.join(filter(str.isdigit, qty_raw)) or '1'
                data['明細'].append({
                    '内容': '■' + str(non_none.get(2, '')).strip(),
                    '数量': int(qty_num),
                    '単価': non_none.get(46, 0),
                })

    return data
```

### Step 3: 請求書テンプレートへの書き込み

openpyxlで読み込んでデータを書き込む。
**テンプレートはRead-Onlyマウントのため直接書き込めない。必ずBytesIOを経由するか、一時ファイルに保存してからWrite toolでoutputsに書き出すこと。**

```python
import io
from datetime import date
from openpyxl import load_workbook

def fill_invoice(template_path, data):
    wb = load_workbook(template_path)
    ws = wb.active

    # 発行日（本日の日付）
    ws['G3'] = date.today()

    # 請求書番号（空欄・手入力）
    ws['G4'] = ''

    # 宛名会社（A5:B5はマージ済み → A5が親セル）
    ws['A5'] = data['宛名会社']

    # 担当者（A6:C6マージ → A6が親セル）
    ws['A6'] = data['担当者'] + '様' if data['担当者'] else ''

    # 件名（A9:C11マージ → A9が親セル）
    ws['A9'] = data['件名']

    # 明細（17行目から最大29件）
    for i, item in enumerate(data['明細']):
        row = 17 + i
        if row > 45:
            break
        ws[f'A{row}'] = item['内容']
        ws[f'D{row}'] = item['数量']
        ws[f'F{row}'] = item['単価']
        # G列の金額はIF数式が自動計算するため入力不要

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
```

BytesIOで得たバイナリは **Write tool** を使って `/sessions/.../mnt/outputs/` に書き込む。

### Step 4: 出力ファイル名の決定

```
【請求書】{件名の最初の40文字}.xlsx
```

ファイル名に使えない文字（`\ / : * ? " < > |`）はアンダースコアに置換する。

### Step 5: 完了メッセージ

ユーザーに以下を伝える：
- 出力ファイルへのリンク（`computer://` リンク）
- 自動入力された項目のサマリー（宛名・担当者・件名・明細件数）
- 請求書番号は空欄のため手入力が必要な旨

## 注意事項

- `openpyxl` が必要。未インストールの場合は `pip install openpyxl --break-system-packages` を実行する
- 明細行は最大29件（行17〜45）まで対応
- テンプレートにはマージセルが多数ある。マージセルの親セル以外に値を書くと `AttributeError: 'MergedCell'` になるため、必ず親セルに書き込むこと
- 金額・消費税・合計はテンプレートの数式（`=IF(D17*F17=0,"",D17*F17)` 等）が自動計算するため入力しない
- 担当者名に全角スペース（\u3000）が含まれる場合は半角スペースに変換する
