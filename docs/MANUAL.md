# FlowBase-VBA ユーザーマニュアル

## 目次

1. [はじめに](#はじめに)
2. [セットアップ](#セットアップ)
3. [各機能の使い方](#各機能の使い方)
4. [シート構成](#シート構成)
5. [トラブルシューティング](#トラブルシューティング)

---

## はじめに

FlowBase-VBA は、Excel VBA のみで動作するプロジェクト・タスク管理ツールです。
外部の実行ファイル（.exe）を必要とせず、マクロ有効ブック（.xlsm）だけで完結します。

### 主な特徴

- **VBA完結型**: 外部ツール不要、Excel だけで動作
- **リアルタイム更新**: ブックを閉じずに即座に反映
- **カスタマイズ可能**: DEF_* シートで各種設定を変更可能

---

## セットアップ

### 1. VBAモジュールのインポート

1. Excelファイル（.xlsm）を開く
2. `Alt + F11` でVBAエディタを開く
3. 「ファイル」→「ファイルのインポート」を選択
4. 以下の順序でモジュールをインポート:

```
【必須】依存関係順にインポート

1. src/vba/common/Mod_Constants.bas   ← 最初にインポート（定数定義）
2. src/vba/common/Utl_Logger.bas
3. src/vba/common/Utl_Table.bas
4. src/vba/common/Utl_Sheet.bas
5. src/vba/common/Utl_File.bas
6. src/vba/manager/Mgr_Logger.cls
7. src/vba/tools/Pst_*.bas            ← 最後にインポート
```

### 2. 必須シートの準備

以下のシートがブックに存在することを確認してください:

| シート名 | 用途 |
|----------|------|
| `DEF_Parameter` | システムパラメータ定義 |
| `DEF_SheetPrefix` | シートソート順定義 |
| `DEF_project_category` | プロジェクトカテゴリ定義 |
| `UI_Index` | シート一覧表示 |
| `UI_ProjectIndex` | プロジェクト一覧表示 |
| `UI_AddSheet` | シート追加用入力フォーム |

### 3. DEF_Parameter の設定

`DEF_Parameter` シートに `Tbl_Start:Parameter` マーカーを設定し、以下のパラメータを定義:

| name | value | 説明 |
|------|-------|------|
| `OBSIDIAN_PATH_FROM_SYSTEM_ROOT` | `C:\Users\...\Obsidian` | Obsidian Vault のパス |
| `PROJECT_SHEET_NAME_BY_CATEGORY_FY_SEQ` | `DEF_PJ-CATEGORY-FY-SEQ` | PJシートテンプレート名 |
| `PERSONAL_TASK_SHEET_TEMPLATE` | `DEF_PT-Name` | PTシートテンプレート名 |

---

## 各機能の使い方

### IndexUpdate（シート一覧更新）

全シートの情報を `UI_Index` シートに反映します。

```vba
' 実行方法
Call Pst_IndexUpdate.IndexUpdate
```

**動作:**
- 全シートの名前、プレフィックス、header_info を収集
- UI_Index の IndexTable テーブルに書き込み

---

### ProjectIndexUpdate（プロジェクト一覧更新）

`PJ-` プレフィックスのシート情報を `UI_ProjectIndex` に反映します。

```vba
Call Pst_ProjectIndexUpdate.ProjectIndexUpdate
```

**動作:**
- PJ-* シートの header_info を収集
- UI_ProjectIndex の project_index テーブルに書き込み

---

### AddProjectSheet（プロジェクトシート作成）

テンプレートから新規プロジェクトシートを作成します。

```vba
Call Pst_AddProjectSheet.AddProjectSheet
```

**事前準備:**
1. `UI_AddSheet` の `AddProjectManagementSheet` テーブルに入力:
   - `project_category`: プロジェクトカテゴリ（DEF_project_category から選択）
   - `summary`: プロジェクト概要
   - その他必要な項目

**動作:**
- カテゴリコードと年度から連番を計算
- テンプレートをコピーして新シート作成（例: `PJ-DEV-FY25-01`）
- header_info を更新

---

### AddPersonalTaskSheet（個人タスクシート作成）

テンプレートから個人タスクシートを作成します。

```vba
Call Pst_AddPersonalTaskSheet.AddPersonalTaskSheet
```

**事前準備:**
1. `UI_AddSheet` の `AddParsonalTaskSheet` テーブルに入力:
   - `owner_name`: 担当者名（必須）

**動作:**
- テンプレートをコピーして新シート作成（例: `PT-Yamada`）

---

### UpdatePersonalTask（個人タスク集約）

`Doing` ステータスのタスクを担当者の個人シートに集約します。

```vba
Call Pst_UpdatePersonalTask.UpdatePersonalTask
```

**動作:**
- 全 PJ-* シートから `Kanban_Status = Doing` のタスクを収集
- 担当者（owner_primary）の PT-* シートに書き込み

---

### UpdateTaskUrgent（期限間近タスク表示）

期限が近いタスクを `OUT_TaskUrgent` シートに表示します。

```vba
Call Pst_UpdateTaskUrgent.UpdateTaskUrgent
```

**動作:**
- 全 PJ-* シートから期限3日以内または期限超過のタスクを収集
- OUT_TaskUrgent の TaskUrgent テーブルに書き込み

---

### SortSheets（シートソート）

`DEF_SheetPrefix` の定義に従ってシートを並び替えます。

```vba
Call Pst_SortSheets.SortSheets
```

**動作:**
- DEF_SheetPrefix からプレフィックス優先度を読み込み
- 全シートを優先度順→名前順にソート

---

### JumpNextUpdate（次の更新シートへジャンプ）

`update_flag = YES` のシートへジャンプします。

```vba
Call Pst_JumpNextUpdate.JumpNextUpdate
```

**動作:**
- 全シートの header_info から update_flag を確認
- YES のシートがあればジャンプ

---

### OutputToObsidian（Obsidian出力）

現在のシートのタスクを Obsidian ノートとして出力します。

```vba
Call Pst_OutputToObsidian.OutputToObsidian
```

**事前準備:**
- `DEF_Parameter` に `OBSIDIAN_PATH_FROM_SYSTEM_ROOT` を設定
- `M_Cov_WBS-Obsidian` シートでフィールドマッピングを定義

**動作:**
- 現在シートの TaskList を読み込み
- YAML フロントマター付きの Markdown ファイルを生成
- 指定フォルダに出力

---

## シート構成

### プレフィックス規則

| プレフィックス | 用途 | 例 |
|----------------|------|-----|
| `UI_` | ユーザーインターフェース | UI_Index, UI_AddSheet |
| `PJ-` | プロジェクト管理 | PJ-DEV-FY25-01 |
| `PT-` | 個人タスク管理 | PT-Yamada |
| `DEF_` | 定義・テンプレート | DEF_Parameter |
| `M_` | マスタデータ | M_Cov_WBS-Obsidian |
| `OUT_` | 出力シート | OUT_TaskUrgent |

### Tbl_Start マーカー

各シートでは `Tbl_Start:<マーカー名>` でテーブル位置を識別します。
マーカーはA列に配置し、次の行がヘッダー行となります。

```
A1: Tbl_Start:header_info
A2: key        | B2: value
A3: project_id | B3: PJ-DEV-FY25-01
A4: summary    | B4: 新機能開発
```

---

## トラブルシューティング

### 「定義されていません」エラー

**原因:** Mod_Constants.bas がインポートされていない

**対処:**
1. VBAエディタで Mod_Constants モジュールが存在するか確認
2. なければ `src/vba/common/Mod_Constants.bas` をインポート

---

### 「ByRef 引数の型が一致しません」エラー

**原因:** 変数の型が関数の引数と合わない

**対処:**
- 最新のモジュールファイルを再インポート
- 古いモジュールを削除してから新しいものをインポート

---

### テンプレートが見つからないエラー

**原因:** DEF_Parameter のテンプレート名が正しくない

**対処:**
1. DEF_Parameter シートを確認
2. `Tbl_Start:Parameter` マーカーが存在するか確認
3. `name` 列と `value` 列のヘッダー名を確認
4. テンプレートシートが実際に存在するか確認

---

### ログファイルの確認

各ツールは `logs/` フォルダにログを出力します:

```
logs/vba_indexupdate_20260202.log
logs/vba_sortsheets_20260202.log
```

エラー発生時はログファイルを確認してください。

---

## バージョン履歴

| Version | Date | Changes |
|---------|------|---------|
| 1.0.0 | 2026-02-02 | 初版リリース（VBA完結版） |

---

## ライセンス

MIT License
