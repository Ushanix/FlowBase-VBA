Option Explicit

' ============================================
' Module   : Mod_Constants
' Layer    : Common
' Purpose  : Centralized constants for FlowBase-VBA
' Version  : 1.0.0
' Created  : 2026-02-02
' Note     : Migrated from Python contracts.py
' ============================================

' ============================================
' Tbl_Start Marker Names
' Used to locate data tables in sheets
' ============================================
Public Const TBL_HEADER_INFO As String = "header_info"
Public Const TBL_ACTIONS As String = "actions"
Public Const TBL_STATUS As String = "status"
Public Const TBL_SCHEDULE As String = "schedule"
Public Const TBL_TASK_LIST As String = "TaskList"
Public Const TBL_PERSONAL_TASK As String = "PersonalTask"
Public Const TBL_TASK_URGENT As String = "TaskUrgent"
Public Const TBL_INDEX_HEADER As String = "index_header"
Public Const TBL_INDEX_TABLE As String = "IndexTable"
Public Const TBL_PROJECT_INDEX As String = "project_index"
Public Const TBL_TASK_INDEX As String = "task_index"
Public Const TBL_ADD_PROJECT_SHEET As String = "AddProjectManagementSheet"
Public Const TBL_ADD_PERSONAL_TASK_SHEET As String = "AddParsonalTaskSheet"
Public Const TBL_PARAMETER As String = "Parameter"

' ============================================
' Sheet Name Prefixes
' Used to identify sheet types
' ============================================
Public Const PREFIX_PROJECT As String = "PJ-"
Public Const PREFIX_PERSONAL As String = "PT-"
Public Const PREFIX_TEMPLATE_PROJECT As String = "DEF_PJ-"
Public Const PREFIX_TEMPLATE_PERSONAL As String = "DEF_PT-"
Public Const PREFIX_MASTER As String = "M_"
Public Const PREFIX_DEFINITION As String = "DEF_"
Public Const PREFIX_UI As String = "UI_"
Public Const PREFIX_OUTPUT As String = "OUT_"

' ============================================
' Fixed Sheet Names
' ============================================
Public Const SHEET_UI_INDEX As String = "UI_Index"
Public Const SHEET_UI_PROJECT_INDEX As String = "UI_ProjectIndex"
Public Const SHEET_UI_ADD_SHEET As String = "UI_AddSheet"
Public Const SHEET_DEF_PARAMETER As String = "DEF_Parameter"
Public Const SHEET_DEF_SHEET_PREFIX As String = "DEF_SheetPrefix"
Public Const SHEET_DEF_PROJECT_CATEGORY As String = "DEF_project_category"
Public Const SHEET_OUT_TASK_URGENT As String = "OUT_TaskUrgent"

' ============================================
' Parameter Keys (from DEF_Parameter sheet)
' ============================================
Public Const PARAM_OBSIDIAN_PATH As String = "OBSIDIAN_PATH_FROM_SYSTEM_ROOT"
Public Const PARAM_PROJECT_TEMPLATE As String = "PROJECT_SHEET_NAME_BY_CATEGORY_FY_SEQ"
Public Const PARAM_PERSONAL_TEMPLATE As String = "PERSONAL_TASK_SHEET_TEMPLATE"
Public Const PARAM_LAST_MTG_DATE As String = "LAST-MTG-DATE"

' ============================================
' Default Values
' ============================================
Public Const DEFAULT_PROJECT_TEMPLATE As String = "DEF_PJ-CATEGORY-FY-SEQ"
Public Const DEFAULT_PERSONAL_TEMPLATE As String = "DEF_PT-Name"
Public Const DEFAULT_SORT_ORDER As Long = 9999
Public Const URGENCY_THRESHOLD_DAYS As Long = 3

' ============================================
' Kanban Status Values
' ============================================
Public Const KANBAN_BACKLOG As String = "Backlog"
Public Const KANBAN_TODO As String = "Todo"
Public Const KANBAN_DOING As String = "Doing"
Public Const KANBAN_REVIEW As String = "Review"
Public Const KANBAN_DONE As String = "Done"
Public Const KANBAN_BLOCKED As String = "Blocked"
Public Const KANBAN_ON_HOLD As String = "OnHold"

' ============================================
' Personal Task Sheet Fixed Values
' ============================================
Public Const FIXED_SHEET_ROLE As String = "personal_queue"

' ============================================
' PersonalTaskTable Headers (column order)
' ============================================
Public Function GetPersonalTaskHeaders() As Variant
    GetPersonalTaskHeaders = Array( _
        "no", "task_id", "src_project_id", "src_sheet_name", _
        "task_name", "description", "owner_primary", "owner_secondary", _
        "Kanban_Status", "MoSCoW_Priority", "story_point", "DoD_check", _
        "start_date", "end_date", "last_update", "update_flag", _
        "feature_key", "sprint_id")
End Function

' ============================================
' TaskUrgent Headers (column order)
' ============================================
Public Function GetTaskUrgentHeaders() As Variant
    GetTaskUrgentHeaders = Array( _
        "no", "src_project_id", "src_sheet_name", "task_id", _
        "task_name", "summary", "owner_primary", "owner_secondary", _
        "Kanban_Status", "MoSCoW_Priority", "story_point", "DoD_check", _
        "start_date", "end_date", "last_update", "update_flag", _
        "feature_key", "sprint_id")
End Function
