# invoice-teams — 注文書Excel → 請求書Excel 自動生成スキル

注文書（Excel）と請求書テンプレート（Excel）をアップロードすることで、請求書を自動生成します。

## できること

- 注文書Excelから以下を自動抽出・転記
  - 発注元会社名、担当者名、注文件名
  - 明細（内容・数量・単価）最大29件
- 金額・消費税・合計はテンプレートの数式が自動計算
- 請求書番号は空欄（手入力）

## インストール方法

```
/mnt/skills/user/invoice-teams/
├── SKILL.md
└── scripts/
    └── extract_po_excel.py
```

## 使い方

Claude に以下の2ファイルをアップロードして話しかけます：

1. 注文書 Excel（`.xlsx`）
2. 請求書テンプレート Excel（`.xlsx`）

発言例：
- 「請求書を作成して」
- 「注文書から請求書に転記して」
- 「自動入力して」

## 必要なライブラリ

```bash
pip install openpyxl --break-system-packages
```

## 対応フォーマット

JR東日本企画フォーマットの注文書に対応しています。  
発注元会社名は「セールス番号／経費番号」行の AU列、担当者名は「担当者」を含む行の AU列から取得します。  
フォーマットが異なる場合は `SKILL.md` および `scripts/extract_po_excel.py` を調整してください。
