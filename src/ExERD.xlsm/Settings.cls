VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
'  Copyright(C) 2005-2009 YAGI Hiroto All Right Reserved
'
'  Licensed under the Apache License, Version 2.0 (the "License");
'  you may not use this file except in compliance with the License.
'  You may obtain a copy of the License at
'
'      http://www.apache.org/licenses/LICENSE-2.0
'
'  Unless required by applicable law or agreed to in writing, software
'  distributed under the License is distributed on an "AS IS" BASIS,
'  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'  See the License for the specific language governing permissions and
'  limitations under the License.
'

'
' ExERDアプリケーションの各種設定を操作するクラス
'
'
'

Private Const POS_INVARID                       As Integer = -1
'********** DEFAULT SETTINGS **********
'DDL DEFAULT SETTINGS
Private Const DEFAULT_START_ROW                 As Integer = 3
Private Const DEFAULT_POS_OBJECT_TYPE           As Integer = 1
Private Const DEFAULT_POS_LOGICAL_TABLENAME     As Integer = 2
Private Const DEFAULT_POS_PHYSICAL_TABLENAME    As Integer = 3
Private Const DEFAULT_POS_COLID                 As Integer = 4
Private Const DEFAULT_POS_LOGICAL_COLNAME       As Integer = 5
Private Const DEFAULT_POS_PHYSICAL_COLNAME      As Integer = 6
Private Const DEFAULT_POS_DATATYPE              As Integer = 7
Private Const DEFAULT_POS_DATALENGTH            As Integer = 8
Private Const DEFAULT_POS_NOTNULL               As Integer = 9
Private Const DEFAULT_POS_PRIMARYKEY            As Integer = 10
Private Const DEFAULT_POS_FOREIGNKEY            As Integer = 11
Private Const DEFAULT_POS_DEPENDENCE_TABLENAME  As Integer = 12
Private Const DEFAULT_POS_RELATION_TYPE         As Integer = 13
Private Const DEFAULT_POS_DEFAULT_VALUE         As Integer = 14

'FORMAT DEFAULT SETTINGS
Private Const DEFAULT_FONTSIZE                  As Integer = 10
Private Const DEFAULT_MARGIN_TOP                As Integer = 80
Private Const DEFAULT_MARGIN_LEFT               As Integer = 80
Private Const DEFAULT_INTERVAL                  As Integer = 20
Private Const DEFAULT_WIDTHLIMIT                As Integer = 800

'DDL FILE SETTINGS
Private Const DEFAULT_DDL_OUTPUT_PATH           As String = ""
Private Const DEFAULT_DDL_OUTPUT_FILE           As String = "ExERD_DDL.sql"
Private Const DEFAULT_DDL_SEP_STR               As String = "/"
Private Const DEFAULT_DDL_COMMENT_STR           As String = "--"
Private Const DEFAULT_DDL_OPEN_NOTEPAD          As Integer = 1
Private Const DEFAULT_DDL_OUTPUT_LOGICAL_NAME   As Integer = 1

'ODBC REVERSE SETTINGS
Private Const DEFAULT_ODBC_DATATYPEFILE         As String = ""

'USER INTERFACE DEFAULT SETTINGS
Private Const DEFAULT_MODEL_MODE                As Integer = 0
Private Const DEFAULT_OUTPUT_RELATION           As Integer = 1
Private Const DEFAULT_OUTPUT_ELISION            As Integer = 0

'********** KEY SETTINGS **********
'DDL KEY SETTINGS
Private Const KEY_START_ROW                     As String = "START_ROW"
Private Const KEY_POS_OBJECT_TYPE               As String = "POS_OBJECT_TYPE"
Private Const KEY_POS_LOGICAL_TABLENAME         As String = "POS_LOGICAL_TABLENAME"
Private Const KEY_POS_PHYSICAL_TABLENAME        As String = "POS_PHYSICAL_TABLENAME"
Private Const KEY_POS_COLID                     As String = "POS_COLID"
Private Const KEY_POS_LOGICAL_COLNAME           As String = "POS_LOGICAL_COLNAME"
Private Const KEY_POS_PHYSICAL_COLNAME          As String = "POS_PHYSICAL_COLNAME"
Private Const KEY_POS_DATATYPE                  As String = "POS_DATATYPE"
Private Const KEY_POS_DATALENGTH                As String = "POS_DATALENGTH"
Private Const KEY_POS_NOTNULL                   As String = "POS_NOTNULL"
Private Const KEY_POS_PRIMARYKEY                As String = "POS_PRIMARYKEY"
Private Const KEY_POS_FOREIGNKEY                As String = "POS_FOREIGNKE"
Private Const KEY_POS_DEPENDENCE_TABLENAME      As String = "POS_DEPENDENCE_TABLENAME"
Private Const KEY_POS_RELATION_TYPE             As String = "POS_RELATION_TYPE"
Private Const KEY_POS_DEFAULT_VALUE             As String = "POS_DEFAULT_VALUE"

'FORMAT KEY SETTINGS
Private Const KEY_FONTSIZE                      As String = "FONT_SIZE"
Private Const KEY_MARGIN_TOP                    As String = "MARGIN_TOP"
Private Const KEY_MARGIN_LEFT                   As String = "MARGIN_LEFT"
Private Const KEY_INTERVAL                      As String = "INTERVAL"
Private Const KEY_WIDTHLIMIT                    As String = "WIDTHLIMIT"

'DDL FILE SETTING
Private Const KEY_DDL_OUTPUT_PATH               As String = "DDL_OUTPUT_PATH"
Private Const KEY_DDL_OUTPUT_FILE               As String = "DDL_OUTPUT_FILE"
Private Const KEY_DDL_SEP_STR                   As String = "DDL_SEP_STR"
Private Const KEY_DDL_COMMENT_STR               As String = "DDL_COMMENT_STR"
Private Const KEY_DDL_OPEN_NOTEPAD              As String = "DDL_OPEN_NOTEPAD"
Private Const KEY_DDL_OUTPUT_LOGICAL_NAME       As String = "DDL_OUTPUT_LOGICAL_NAME"

'ODBC REVERSE KEY SETTINGS
Private Const KEY_ODBC_DATATYPEFILE             As String = "ODBC_DATATYPEFILE"
    
'USER INTERFACE KEY SETTINGS
Private Const KEY_MODEL_MODE                    As String = "MODEL_MODE"
Private Const KEY_OUTPUT_RELATION               As String = "OUTPUT_RELATION"
Private Const KEY_OUTPUT_ELISION                As String = "OUTPUT_ELISION"

'DDL MEMBERS
Private m_Start_Row                     As Integer
Private m_Colpos_ObjectType             As Integer
Private m_Colpos_LogicalTableName       As Integer
Private m_Colpos_PhysicalTableName      As Integer
Private m_Colpos_ColId                  As Integer
Private m_Colpos_LogicaColName          As Integer
Private m_Colpos_PhysicallColName       As Integer
Private m_Colpos_DataType               As Integer
Private m_Colpos_DataLength             As Integer
Private m_Colpos_NotNull                As Integer
Private m_Colpos_PrimaryKey             As Integer
Private m_Colpos_Foreignkey             As Integer
Private m_Colpos_DependenceTableName    As Integer
Private m_Colpos_RelationType           As Integer
Private m_Colpos_DefaultValue           As Integer

'FONT MEMBERS
Private m_FontSize                      As Integer
Private m_MarginTop                     As Integer
Private m_MarginLeft                    As Integer
Private m_Interval                      As Integer
Private m_Widthlimit                    As Integer

'DDL FILE
Private m_DDLOutputPath                 As String
Private m_DDLOutputFile                 As String
Private m_DDLSeqStr                     As String
Private m_DDLCommentStr                 As String
Private m_DDLOpenNotepad                As Integer
Private m_DDLOutputLogicalName          As Integer

'ODBC REVERSE
Private m_Odbc_DatatypeFile             As String

'USER INTERFACE
Private m_ModelMode                     As Integer
Private m_OutputRelation                As Integer
Private m_OutputElision                 As Integer

'
'
'
'
Public Function loadDDLSettings(ByRef book As Excel.Workbook, isUseDefaultPos As Boolean) As Long
    
    Call setStartRow(getAppSetting(book, _
        KEY_START_ROW, _
        IIf(isUseDefaultPos, DEFAULT_START_ROW, POS_INVARID)))
    Call setColposObjectType(getAppSetting(book, _
        KEY_POS_OBJECT_TYPE, _
        IIf(isUseDefaultPos, DEFAULT_POS_OBJECT_TYPE, POS_INVARID)))
    Call setColposLogicalTableName(getAppSetting(book, _
        KEY_POS_LOGICAL_TABLENAME, _
        IIf(isUseDefaultPos, DEFAULT_POS_LOGICAL_TABLENAME, POS_INVARID)))
    Call setColposPhysicalTableName(getAppSetting(book, _
        KEY_POS_PHYSICAL_TABLENAME, _
        IIf(isUseDefaultPos, DEFAULT_POS_PHYSICAL_TABLENAME, POS_INVARID)))
    Call setColposColId(getAppSetting(book, _
        KEY_POS_COLID, IIf(isUseDefaultPos, _
        DEFAULT_POS_COLID, POS_INVARID)))
    Call setColposLogicaColName(getAppSetting(book, _
        KEY_POS_LOGICAL_COLNAME, _
        IIf(isUseDefaultPos, DEFAULT_POS_LOGICAL_COLNAME, POS_INVARID)))
    Call setColposPhysicalColName(getAppSetting(book, _
        KEY_POS_PHYSICAL_COLNAME, _
        IIf(isUseDefaultPos, DEFAULT_POS_PHYSICAL_COLNAME, POS_INVARID)))
    Call setColposDataType(getAppSetting(book, _
        KEY_POS_DATATYPE, _
        IIf(isUseDefaultPos, DEFAULT_POS_DATATYPE, POS_INVARID)))
    Call setColposDataLength(getAppSetting(book, _
        KEY_POS_DATALENGTH, _
        IIf(isUseDefaultPos, DEFAULT_POS_DATALENGTH, POS_INVARID)))
    Call setColposNotNull(getAppSetting(book, _
        KEY_POS_NOTNULL, _
        IIf(isUseDefaultPos, DEFAULT_POS_NOTNULL, POS_INVARID)))
    Call setColposPrimaryKey(getAppSetting(book, _
        KEY_POS_PRIMARYKEY, _
        IIf(isUseDefaultPos, DEFAULT_POS_PRIMARYKEY, POS_INVARID)))
    Call setColposForeingKey(getAppSetting(book, _
        KEY_POS_FOREIGNKEY, _
        IIf(isUseDefaultPos, DEFAULT_POS_FOREIGNKEY, POS_INVARID)))
    Call setColposDependenceTableName(getAppSetting(book, _
        KEY_POS_DEPENDENCE_TABLENAME, _
        IIf(isUseDefaultPos, DEFAULT_POS_DEPENDENCE_TABLENAME, POS_INVARID)))
    Call setColposRelationType(getAppSetting(book, _
        KEY_POS_RELATION_TYPE, _
        IIf(isUseDefaultPos, DEFAULT_POS_RELATION_TYPE, POS_INVARID)))
    Call setColposDefaultValue(getAppSetting(book, _
        KEY_POS_DEFAULT_VALUE, _
        IIf(isUseDefaultPos, DEFAULT_POS_DEFAULT_VALUE, POS_INVARID)))
        
End Function
'
'
'
Public Function loadFormatSettings(ByRef book As Excel.Workbook, isUseDefaultPos As Boolean) As Long
    
    Call setFontSize(getAppSetting(book, _
        KEY_FONTSIZE, _
        IIf(isUseDefaultPos, DEFAULT_FONTSIZE, POS_INVARID)))
    Call setMarginTop(getAppSetting(book, _
        KEY_MARGIN_TOP, _
        IIf(isUseDefaultPos, DEFAULT_MARGIN_TOP, POS_INVARID)))
    Call setMarginLeft(getAppSetting(book, _
        KEY_MARGIN_LEFT, _
        IIf(isUseDefaultPos, DEFAULT_MARGIN_LEFT, POS_INVARID)))
    Call setInterval(getAppSetting(book, _
        KEY_INTERVAL, _
        IIf(isUseDefaultPos, DEFAULT_INTERVAL, POS_INVARID)))
    Call setWidthlimit(getAppSetting(book, _
        KEY_WIDTHLIMIT, _
        IIf(isUseDefaultPos, DEFAULT_WIDTHLIMIT, POS_INVARID)))
        
End Function
'
'
'
Public Function loadDDLOutputSettings(ByRef book As Excel.Workbook, isUseDefault As Boolean) As Long
    
    Call setDDLOutputPath(getAppSetting(book, _
        KEY_DDL_OUTPUT_PATH, _
        DEFAULT_DDL_OUTPUT_PATH))
    Call setDDLOutputFile(getAppSetting(book, _
        KEY_DDL_OUTPUT_FILE, _
        DEFAULT_DDL_OUTPUT_FILE))
    Call setDDLSepStr(getAppSetting(book, _
        KEY_DDL_SEP_STR, _
        DEFAULT_DDL_SEP_STR))
    Call setDDLCommentStr(getAppSetting(book, _
        KEY_DDL_COMMENT_STR, _
        DEFAULT_DDL_COMMENT_STR))
    Call setDDLOpenNodepad(getAppSetting(book, _
        KEY_DDL_OPEN_NOTEPAD, _
        DEFAULT_DDL_OPEN_NOTEPAD))
    Call setDDLOutputLogicalName(getAppSetting(book, _
        KEY_DDL_OUTPUT_LOGICAL_NAME, _
        DEFAULT_DDL_OUTPUT_LOGICAL_NAME))
    
End Function
'
'
'
Public Function loadOdbcSettings(ByRef book As Excel.Workbook, isUseDefaultPos As Boolean) As Long
    
    Call setOdbcDatatypeFile(getAppSetting(book, _
        KEY_ODBC_DATATYPEFILE, _
        DEFAULT_ODBC_DATATYPEFILE))
    
End Function
'
'
'
Public Function loadUISettings(ByRef book As Excel.Workbook, isUseDefaultPos As Boolean) As Long
    
    Call setModelMode(getAppSetting(book, _
        KEY_MODEL_MODE, _
        DEFAULT_MODEL_MODE))
    Call setOutputRelation(getAppSetting(book, _
        KEY_OUTPUT_RELATION, _
        DEFAULT_OUTPUT_RELATION))
    Call setOutputElision(getAppSetting(book, _
        KEY_OUTPUT_ELISION, _
        DEFAULT_OUTPUT_ELISION))
    
End Function
'
' すべての設定の読み込み
'
Public Function loadSettings(ByRef book As Excel.Workbook, isUseDefault As Boolean) As Long
    
    Call loadDDLSettings(book, isUseDefault)
    Call loadFormatSettings(book, isUseDefault)
    Call loadDDLOutputSettings(book, isUseDefault)
    Call loadOdbcSettings(book, isUseDefault)
    Call loadUISettings(book, isUseDefault)
    
End Function
'
' DDL デフォルト設定値の読み込み
'
Public Function loadDDLDefaultSettings() As Long
    
    Call setStartRow(DEFAULT_START_ROW)
    Call setColposObjectType(DEFAULT_POS_OBJECT_TYPE)
    Call setColposLogicalTableName(DEFAULT_POS_LOGICAL_TABLENAME)
    Call setColposPhysicalTableName(DEFAULT_POS_PHYSICAL_TABLENAME)
    Call setColposColId(DEFAULT_POS_COLID)
    Call setColposLogicaColName(DEFAULT_POS_LOGICAL_COLNAME)
    Call setColposPhysicalColName(DEFAULT_POS_PHYSICAL_COLNAME)
    Call setColposDataType(DEFAULT_POS_DATATYPE)
    Call setColposDataLength(DEFAULT_POS_DATALENGTH)
    Call setColposNotNull(DEFAULT_POS_NOTNULL)
    Call setColposPrimaryKey(DEFAULT_POS_PRIMARYKEY)
    Call setColposForeingKey(DEFAULT_POS_FOREIGNKEY)
    Call setColposDependenceTableName(DEFAULT_POS_DEPENDENCE_TABLENAME)
    Call setColposRelationType(DEFAULT_POS_RELATION_TYPE)
    Call setColposDefaultValue(DEFAULT_POS_DEFAULT_VALUE)
    
End Function
'
'
'
Public Function loadFormatDefaultSettings() As Long
    
    Call setFontSize(DEFAULT_FONTSIZE)
    Call setMarginTop(DEFAULT_MARGIN_TOP)
    Call setMarginLeft(DEFAULT_MARGIN_LEFT)
    Call setInterval(DEFAULT_INTERVAL)
    Call setWidthlimit(DEFAULT_WIDTHLIMIT)
    
End Function
'
'
'
Public Function loadDDLOutputDefaultSettings() As Long
    
    Call setDDLOutputPath(DEFAULT_DDL_OUTPUT_PATH)
    Call setDDLOutputFile(DEFAULT_DDL_OUTPUT_FILE)
    Call setDDLSepStr(DEFAULT_DDL_SEP_STR)
    Call setDDLCommentStr(DEFAULT_DDL_COMMENT_STR)
    Call setDDLOpenNodepad(DEFAULT_DDL_OPEN_NOTEPAD)
    Call setDDLOutputLogicalName(DEFAULT_DDL_OUTPUT_LOGICAL_NAME)
    
End Function
'
'
'
Public Function loadODBCDefaultSettings() As Long

    Call setOdbcDatatypeFile(DEFAULT_ODBC_DATATYPEFILE)
    
End Function
'
'
'
Public Function loadUIDefaultSettings() As Long

    Call setModelMode(DEFAULT_MODEL_MODE)
    Call setOutputRelation(DEFAULT_OUTPUT_RELATION)
    Call setOutputElision(DEFAULT_OUTPUT_ELISION)
    
End Function
'
'
'
Public Function loadDefaultSettings() As Long
    
    Call loadDDLDefaultSettings
    Call loadFormatDefaultSettings
    Call loadUIDefaultSettings
    
End Function
'
'
'
Public Function saveDDLSettings(ByRef book As Excel.Workbook) As Long
    
    Call setAppSetting(book, KEY_START_ROW, getStartRow())
    Call setAppSetting(book, KEY_POS_OBJECT_TYPE, getColposObjectType())
    Call setAppSetting(book, KEY_POS_LOGICAL_TABLENAME, getColposLogicalTableName())
    Call setAppSetting(book, KEY_POS_PHYSICAL_TABLENAME, getColposPhysicalTableName())
    Call setAppSetting(book, KEY_POS_COLID, getColposColId())
    Call setAppSetting(book, KEY_POS_LOGICAL_COLNAME, getColposLogicaColName())
    Call setAppSetting(book, KEY_POS_PHYSICAL_COLNAME, getColposPhysicalColName())
    Call setAppSetting(book, KEY_POS_DATATYPE, getColposDataType())
    Call setAppSetting(book, KEY_POS_DATALENGTH, getColposDataLength())
    Call setAppSetting(book, KEY_POS_NOTNULL, getColposNotNull())
    Call setAppSetting(book, KEY_POS_PRIMARYKEY, getColposPrimaryKey())
    Call setAppSetting(book, KEY_POS_FOREIGNKEY, getColposForeingKey())
    Call setAppSetting(book, KEY_POS_DEPENDENCE_TABLENAME, getColposDependenceTableName())
    Call setAppSetting(book, KEY_POS_RELATION_TYPE, getColposRelationType())
    Call setAppSetting(book, KEY_POS_DEFAULT_VALUE, getColposDefaultValue())
    
End Function
'
'
'
Public Function saveFormatSettings(ByRef book As Excel.Workbook) As Long
    
    Call setAppSetting(book, KEY_FONTSIZE, getFontSize())
    Call setAppSetting(book, KEY_MARGIN_TOP, getMarginTop())
    Call setAppSetting(book, KEY_MARGIN_LEFT, getMarginLeft())
    Call setAppSetting(book, KEY_INTERVAL, getInterval())
    Call setAppSetting(book, KEY_WIDTHLIMIT, getWidthlimit())

End Function
'
'
'
Public Function saveDDLOutputSettings(ByRef book As Excel.Workbook) As Long

    Call setAppSetting(book, KEY_DDL_OUTPUT_PATH, getDDLOutputPath())
    Call setAppSetting(book, KEY_DDL_OUTPUT_FILE, getDDLOutputFile())
    Call setAppSetting(book, KEY_DDL_SEP_STR, getDDLSeqStr())
    Call setAppSetting(book, KEY_DDL_COMMENT_STR, getDDLCommentStr())
    Call setAppSetting(book, KEY_OUTPUT_RELATION, getOutputRelation())
    Call setAppSetting(book, KEY_DDL_OPEN_NOTEPAD, getDDLOpenNodepad())
    Call setAppSetting(book, KEY_DDL_OUTPUT_LOGICAL_NAME, getDDLOutputLogicalName())
    
End Function
'
'
'
Public Function saveOdbcSettings(ByRef book As Excel.Workbook) As Long

    Call setAppSetting(book, KEY_ODBC_DATATYPEFILE, getOdbcDatatypeFile())
    
End Function

'
'
'
Public Function saveUISettings(ByRef book As Excel.Workbook) As Long

    Call setAppSetting(book, KEY_MODEL_MODE, getModelMode())
    Call setAppSetting(book, KEY_OUTPUT_RELATION, getOutputRelation())
    Call setAppSetting(book, KEY_OUTPUT_ELISION, getOutputElision())
    
End Function

'
' 設定の書き込み
'
Public Function saveSettings(ByRef book As Excel.Workbook) As Long
    
    Call saveDDLSettings(book)
    Call saveFormatSettings(book)
    Call saveDDLOutputSettings(book)
    Call saveOdbcSettings(book)
    Call saveUISettings(book)
    
End Function
'
' ========== アクセサ ==========
'
' DDL
Public Sub setStartRow(newVal As Integer)
    m_Start_Row = newVal
End Sub
Public Function getStartRow() As Integer
    getStartRow = m_Start_Row
End Function
Public Sub setColposObjectType(newVal As Integer)
    m_Colpos_ObjectType = newVal
End Sub
Public Function getColposObjectType() As Integer
    getColposObjectType = m_Colpos_ObjectType
End Function
Public Sub setColposLogicalTableName(newVal As Integer)
    m_Colpos_LogicalTableName = newVal
End Sub
Public Function getColposLogicalTableName() As Integer
    getColposLogicalTableName = m_Colpos_LogicalTableName
End Function
Public Sub setColposPhysicalTableName(newVal As Integer)
    m_Colpos_PhysicalTableName = newVal
End Sub
Public Function getColposPhysicalTableName() As Integer
    getColposPhysicalTableName = m_Colpos_PhysicalTableName
End Function
Public Sub setColposColId(newVal As Integer)
    m_Colpos_ColId = newVal
End Sub
Public Function getColposColId() As Integer
    getColposColId = m_Colpos_ColId
End Function
Public Sub setColposLogicaColName(newVal As Integer)
    m_Colpos_LogicaColName = newVal
End Sub
Public Function getColposLogicaColName() As Integer
    getColposLogicaColName = m_Colpos_LogicaColName
End Function
Public Sub setColposPhysicalColName(newVal As Integer)
    m_Colpos_PhysicallColName = newVal
End Sub
Public Function getColposPhysicalColName() As Integer
    getColposPhysicalColName = m_Colpos_PhysicallColName
End Function
Public Sub setColposDataType(newVal As Integer)
    m_Colpos_DataType = newVal
End Sub
Public Function getColposDataType() As Integer
    getColposDataType = m_Colpos_DataType
End Function
Public Sub setColposDataLength(newVal As Integer)
    m_Colpos_DataLength = newVal
End Sub
Public Function getColposDataLength() As Integer
    getColposDataLength = m_Colpos_DataLength
End Function
Public Sub setColposNotNull(newVal As Integer)
    m_Colpos_NotNull = newVal
End Sub
Public Function getColposNotNull() As Integer
    getColposNotNull = m_Colpos_NotNull
End Function
Public Sub setColposPrimaryKey(newVal As Integer)
    m_Colpos_PrimaryKey = newVal
End Sub
Public Function getColposPrimaryKey() As Integer
    getColposPrimaryKey = m_Colpos_PrimaryKey
End Function
Public Sub setColposForeingKey(newVal As Integer)
    m_Colpos_Foreignkey = newVal
End Sub
Public Function getColposForeingKey() As Integer
    getColposForeingKey = m_Colpos_Foreignkey
End Function
Public Sub setColposDependenceTableName(newVal As Integer)
    m_Colpos_DependenceTableName = newVal
End Sub
Public Function getColposDependenceTableName() As Integer
    getColposDependenceTableName = m_Colpos_DependenceTableName
End Function
Public Sub setColposRelationType(newVal As Integer)
    m_Colpos_RelationType = newVal
End Sub
Public Function getColposRelationType() As Integer
    getColposRelationType = m_Colpos_RelationType
End Function
Public Sub setColposDefaultValue(newVal As Integer)
    m_Colpos_DefaultValue = newVal
End Sub
Public Function getColposDefaultValue() As Integer
    getColposDefaultValue = m_Colpos_DefaultValue
End Function
' FORMAT
Public Sub setFontSize(newVal As Integer)
    m_FontSize = newVal
End Sub
Public Function getFontSize() As Integer
    getFontSize = m_FontSize
End Function
Public Sub setMarginTop(newVal As Integer)
    m_MarginTop = newVal
End Sub
Public Function getMarginTop() As Integer
    getMarginTop = m_MarginTop
End Function
Public Sub setMarginLeft(newVal As Integer)
    m_MarginLeft = newVal
End Sub
Public Function getMarginLeft() As Integer
    getMarginLeft = m_MarginLeft
End Function
Public Sub setInterval(newVal As Integer)
    m_Interval = newVal
End Sub
Public Function getInterval() As Integer
    getInterval = m_Interval
End Function
Public Sub setWidthlimit(newVal As Integer)
    m_Widthlimit = newVal
End Sub
Public Function getWidthlimit() As Integer
    getWidthlimit = m_Widthlimit
End Function
'DDL OUTPUT
Public Sub setDDLOutputPath(newVal As String)
    m_DDLOutputPath = newVal
End Sub
Public Function getDDLOutputPath() As String
    getDDLOutputPath = m_DDLOutputPath
End Function
Public Sub setDDLOutputFile(newVal As String)
    m_DDLOutputFile = newVal
End Sub
Public Function getDDLOutputFile() As String
    getDDLOutputFile = m_DDLOutputFile
End Function
Public Sub setDDLSepStr(newVal As String)
    m_DDLSeqStr = newVal
End Sub
Public Function getDDLSeqStr() As String
    getDDLSeqStr = m_DDLSeqStr
End Function
Public Sub setDDLCommentStr(newVal As String)
    m_DDLCommentStr = newVal
End Sub
Public Function getDDLCommentStr() As String
    getDDLCommentStr = m_DDLCommentStr
End Function

'add tatsuo: ddl open as notepad
Public Sub setDDLOpenNodepad(newVal As Integer)
    m_DDLOpenNotepad = newVal
End Sub
Public Function getDDLOpenNodepad() As Integer
    getDDLOpenNodepad = m_DDLOpenNotepad
End Function

'add tatsuo: ddl out logical name as comment
Public Sub setDDLOutputLogicalName(newVal As Integer)
    m_DDLOutputLogicalName = newVal
End Sub
Public Function getDDLOutputLogicalName() As Integer
    getDDLOutputLogicalName = m_DDLOutputLogicalName
End Function


'ODBC
Public Sub setOdbcDatatypeFile(newVal As String)
    m_Odbc_DatatypeFile = newVal
End Sub
Public Function getOdbcDatatypeFile()
    getOdbcDatatypeFile = m_Odbc_DatatypeFile
End Function
'USER INTERFACE
Public Sub setModelMode(newVal As Integer)
    m_ModelMode = newVal
End Sub
Public Function getModelMode() As Integer
    getModelMode = m_ModelMode
End Function
Public Sub setOutputRelation(newVal As Integer)
    m_OutputRelation = newVal
End Sub
Public Function getOutputRelation() As Integer
    getOutputRelation = m_OutputRelation
End Function
Public Sub setOutputElision(newVal As Integer)
    m_OutputElision = newVal
End Sub
Public Function getOutputElision() As Integer
    getOutputElision = m_OutputElision
End Function


