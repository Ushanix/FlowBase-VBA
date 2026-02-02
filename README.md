# FlowBase-VBA

**FlowBase-VBA** は、Excel VBA で完結するプロジェクト・タスク管理ツールです。

プロジェクト内のあらゆる「作業（Work）」を Workflow として一元管理し、
WBS・MTG・Agile・個人タスクなどの手法や形式に依存せず、
「作業がどの状態にあり、次に何をすべきか」を明確にします。

---

## Features

- **IndexUpdate** - 全シート一覧をUI_Indexシートに反映
- **ProjectIndexUpdate** - PJ_*シートの情報をUI_ProjectIndexに更新
- **AddProjectSheet** - テンプレートから新規プロジェクトシート作成
- **AddPersonalTaskSheet** - テンプレートから個人タスクシート作成
- **UpdatePersonalTask** - Doingステータスのタスクを個人シートに集約
- **UpdateTaskUrgent** - 期限間近タスクをOUT_TaskUrgentに表示
- **SortSheets** - プレフィックス優先度でシートをソート
- **JumpNextUpdate** - update_flag=YESの次シートへジャンプ
- **OutputToObsidian** - タスクをObsidianノートとして出力

---

## Directory Structure

```
FlowBase-VBA/
├── .gitignore
├── LICENSE
├── README.md
├── FlowBase-VBA_1.0.0.xlsm
└── src/
    └── vba/
        ├── common/
        │   ├── Mod_Constants.bas   # 定数定義
        │   ├── Utl_Table.bas       # テーブル操作
        │   ├── Utl_Sheet.bas       # シート操作
        │   ├── Utl_File.bas        # ファイル操作（FSO）
        │   └── Utl_Logger.bas      # ログユーティリティ
        ├── manager/
        │   └── Mgr_Logger.cls      # ロガークラス
        └── tools/
            ├── Pst_IndexUpdate.bas
            ├── Pst_ProjectIndexUpdate.bas
            ├── Pst_AddProjectSheet.bas
            ├── Pst_AddPersonalTaskSheet.bas
            ├── Pst_UpdatePersonalTask.bas
            ├── Pst_UpdateTaskUrgent.bas
            ├── Pst_SortSheets.bas
            ├── Pst_JumpNextUpdate.bas
            └── Pst_OutputToObsidian.bas
```

---

## VBA Module Import

VBAモジュールをExcelにインポートする手順:

1. Excelファイル（.xlsm）を開く
2. Alt+F11 でVBAエディタを開く
3. 「ファイル」→「ファイルのインポート」
4. `src/vba/` 以下の `.bas` / `.cls` ファイルを選択

### Import Order (依存関係順)

1. Common modules:
   - `Mod_Constants.bas`
   - `Utl_Logger.bas`
   - `Utl_Table.bas`
   - `Utl_Sheet.bas`
   - `Utl_File.bas`

2. Manager:
   - `Mgr_Logger.cls`

3. Tools (任意の順序):
   - `Pst_*.bas` ファイル

---

## Sheet Naming Convention

| Prefix | Role |
|--------|------|
| `UI_` | 操作画面・メニュー |
| `PJ-` | プロジェクト管理シート |
| `PT-` | 個人タスク管理シート |
| `DEF_` | 定義シート（テンプレート含む） |
| `M_` | マスタシート |
| `OUT_` | 出力シート |

---

## Key Markers

各シートでは `Tbl_Start:<マーカー名>` でテーブル領域を識別します:

| Marker | Description |
|--------|-------------|
| `header_info` | シートヘッダー情報（key-value形式） |
| `TaskList` | タスク一覧テーブル |
| `IndexTable` | UI_Indexの全シート一覧 |
| `project_index` | UI_ProjectIndexのプロジェクト一覧 |
| `PersonalTask` | 個人タスク一覧 |
| `TaskUrgent` | 期限間近タスク一覧 |

---

## Usage

### From Excel Buttons

各機能はシート上のボタンからマクロを呼び出します:

```vba
' 例: IndexUpdate ボタン
Sub ButtonIndexUpdate_Click()
    IndexUpdate
End Sub
```

### Direct VBA Call

```vba
' シートインデックス更新
Call Pst_IndexUpdate.IndexUpdate

' シートソート
Call Pst_SortSheets.SortSheets

' 次の更新シートへジャンプ
Call Pst_JumpNextUpdate.JumpNextUpdate
```

---

## Configuration

### DEF_Parameter Sheet

| Parameter | Description |
|-----------|-------------|
| `OBSIDIAN_PATH_FROM_SYSTEM_ROOT` | Obsidian Vaultのルートパス |
| `PROJECT_SHEET_NAME_BY_CATEGORY_FY_SEQ` | プロジェクトシートテンプレート名 |
| `PERSONAL_TASK_SHEET_TEMPLATE` | 個人タスクシートテンプレート名 |

### DEF_SheetPrefix Sheet

シートソート順序の定義:

| sheet_prefix | sort_order |
|--------------|------------|
| `UI_` | 100 |
| `PJ-` | 200 |
| `PT-` | 300 |
| `OUT_` | 400 |
| `DEF_` | 900 |
| `M_` | 950 |

---

## Kanban Status

| Status | Description |
|--------|-------------|
| `Backlog` | バックログ |
| `Todo` | 着手予定 |
| `Doing` | 作業中 |
| `Review` | レビュー中 |
| `Done` | 完了 |
| `Blocked` | ブロック中 |
| `OnHold` | 保留 |

---

## Documentation

詳細なマニュアルは [docs/MANUAL.md](docs/MANUAL.md) を参照してください。

---

## License

MIT License

---

## Notes

- 本ツールは Excel VBA のみで動作し、外部実行ファイル（.exe）は不要です
- Python 版からの移行により、ブックを閉じる必要がなくなりました
- Obsidian 連携には `M_Cov_WBS-Obsidian` シートでフィールドマッピングを定義します

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0.0 | 2026-02-02 | VBA完結版として再構成 |
