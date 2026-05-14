# Claude Custom Skills

Claude.ai で使えるカスタムスキル集です。  
各スキルの `SKILL.md` を `/mnt/skills/user/<スキル名>/` に配置することで利用できます。

## スキル一覧

| スキル名 | 概要 |
|---|---|
| [invoice-lag](./invoice-lag/) | 発注書PDF → 請求書Excel 自動生成 |
| [invoice-teams](./invoice-teams/) | 注文書Excel → 請求書Excel 自動生成 |

---

## 使い方

### 1. スキルファイルの配置

```
/mnt/skills/user/
├── invoice-lag/
│   └── SKILL.md
└── invoice-teams/
    ├── SKILL.md
    └── scripts/
        └── extract_po_excel.py
```

### 2. Claude に話しかける

スキルが配置されると、以下のような発言で自動的に呼び出されます：

- 「発注書から請求書を作成して」
- 「注文書から請求書に転記して」
- 「自動入力して」

---

## 必要な Python ライブラリ

```bash
pip install pdfplumber openpyxl --break-system-packages
```
