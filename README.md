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

ZIP を解凍して、以下のパスにフォルダごとコピーしてください。

#### Mac の場合

```
~/Library/Application Support/Claude/claude-skills/user/
├── invoice-lag/
│   └── SKILL.md
└── invoice-teams/
    ├── SKILL.md
    └── scripts/
        └── extract_po_excel.py
```

Finder でフォルダを開く場合は、Finder を起動して **移動 → フォルダへ移動**（`Shift + Cmd + G`）から以下を入力してください：
```
~/Library/Application Support/Claude/claude-skills/user/
```

#### Windows の場合

```
C:\Users\<ユーザー名>\AppData\Roaming\Claude\claude-skills\user\
├── invoice-lag\
│   └── SKILL.md
└── invoice-teams\
    ├── SKILL.md
    └── scripts\
        └── extract_po_excel.py
```

エクスプローラーでフォルダを開く場合は、アドレスバーに以下を入力してください：
```
%APPDATA%\Claude\claude-skills\user\
```

### 2. ファイルをアップロードする

Claude のチャット画面でファイルを添付します。

1. チャット入力欄の左側にある 📎 **クリップアイコン**をクリック
2. 対象ファイルを選択してアップロード
   - `invoice-lag`：**発注書 PDF** を1ファイル
   - `invoice-teams`：**注文書 Excel（.xlsx）** と **請求書テンプレート Excel（.xlsx）** の2ファイル

### 3. スキルを呼び出す

ファイルをアップロードしたら、以下のいずれかの方法で呼び出します。

**方法① チャットで話しかける**
```
発注書から請求書を作成して
注文書から請求書に転記して
自動入力して
```

**方法② スラッシュコマンドで呼び出す**

チャット入力欄に `/` を入力するとコマンド一覧が表示されます。
```
/invoice-lag
/invoice-teams
```
該当するスキル名を選択すると自動的に実行されます。

---

## 必要な Python ライブラリ

```bash
pip install pdfplumber openpyxl --break-system-packages
```
