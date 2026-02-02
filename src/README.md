# FlowBase-VBA Source Code

VBA モジュールのソースコード格納ディレクトリ。

## Directory Structure

```
src/
└── vba/
    ├── common/         # 共通ユーティリティモジュール
    │   ├── Mod_Constants.bas
    │   ├── Utl_Table.bas
    │   ├── Utl_Sheet.bas
    │   ├── Utl_File.bas
    │   └── Utl_Logger.bas
    ├── manager/        # マネージャークラス
    │   └── Mgr_Logger.cls
    └── tools/          # ツールモジュール（Presentation層）
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

## Module Naming Convention

| Prefix | Layer | Description |
|--------|-------|-------------|
| `Mod_` | Common | 定数・列挙型定義 |
| `Utl_` | Utility | ユーティリティ関数 |
| `Mgr_` | Manager | ステートフルなマネージャークラス |
| `Pst_` | Presentation | UI/ボタンから呼び出されるエントリポイント |

## Importing Modules

1. Excel VBA エディタ (Alt+F11) を開く
2. 「ファイル」→「ファイルのインポート」を選択
3. **依存関係順にモジュールをインポート（順序重要）**:

```
1. Mod_Constants.bas  ← 最初（他モジュールが依存）
2. Utl_Logger.bas
3. Utl_Table.bas
4. Utl_Sheet.bas
5. Utl_File.bas
6. Mgr_Logger.cls
7. Pst_*.bas          ← 最後
```

**注意:** Mod_Constants.bas を最初にインポートしないと、定数未定義エラーが発生します。

## Dependencies

```
Pst_* (tools)
  ├── Mod_Constants (common)
  ├── Utl_Table (common)
  ├── Utl_Sheet (common)
  ├── Utl_File (common) [OutputToObsidian only]
  └── Utl_Logger (common)

Mgr_Logger (manager)
  └── Utl_Logger (common)
```
