VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExERDDocument"
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
' ExERDアプリケーションのドキュメント操作用クラス
'
'
'

Private m_Tables()  As ErTable
Private m_TableId   As Integer
Public m_Settings  As Settings
Private m_Book      As Excel.Workbook
'
'
'
Public Function loadTableData(ByRef sheet As Excel.Worksheet) As Integer
    Dim col As ErColumun
    Dim posObjectType   As Integer
    Dim posLogTName     As Integer
    Dim posPhyTName     As Integer
    Dim posColId        As Integer
    Dim posLogCName     As Integer
    Dim posPhyCName     As Integer
    Dim posDDtype       As Integer
    Dim posDLen         As Integer
    Dim posNNull        As Integer
    Dim posPK           As Integer
    Dim posFk           As Integer
    Dim posDepend       As Integer
    Dim posRelType      As Integer
    Dim posDefaultVal   As Integer
    
    Dim iRow As Integer
    
    ' ============
    '   Logの準備
    ' ============
    Call Log.initialLog(Util.getLogFilename(ThisWorkbook))
    Call Log.info(">>>>> データの読み込みを開始します")
    
    loadTableData = 0
    ReDim m_Tables(0)
    Call m_Settings.loadSettings(getWorkBook(), False)
    
    posObjectType = m_Settings.getColposObjectType()
    posLogTName = m_Settings.getColposLogicalTableName()
    posPhyTName = m_Settings.getColposPhysicalTableName()
    posColId = m_Settings.getColposColId()
    posLogCName = m_Settings.getColposLogicaColName()
    posPhyCName = m_Settings.getColposPhysicalColName()
    posDDtype = m_Settings.getColposDataType()
    posDLen = m_Settings.getColposDataLength()
    posNNull = m_Settings.getColposNotNull()
    posPK = m_Settings.getColposPrimaryKey()
    posFk = m_Settings.getColposForeingKey()
    posDepend = m_Settings.getColposDependenceTableName()
    posRelType = m_Settings.getColposRelationType()
    posDefaultVal = m_Settings.getColposDefaultValue
    
    If m_Settings.getStartRow < 1 Then
        Exit Function
    End If
    
    Dim strLogTName As String
    Dim strPhyTName As String
    Dim tmpLogTName As String
    Dim tmpPhyTName As String
    Dim tmpObjType  As String
    Dim tmpDepTName As String
    Dim tmpDepCName As String
    Dim tmpIsDep    As Boolean
    
    m_TableId = -1
    iRow = m_Settings.getStartRow
    
    Do
        Set col = New ErColumun
        
        strLogTName = getCells(sheet, iRow, posLogTName)
        strPhyTName = getCells(sheet, iRow, posPhyTName)
        
        Call col.setPos(getCellsAsInt(sheet, iRow, posColId))
        Call col.setLogicName(getCells(sheet, iRow, posLogCName))
        Call col.setPhysicName(getCells(sheet, iRow, posPhyCName))
        Call col.setDataType(getCells(sheet, iRow, posDDtype))
        Call col.setDataLength(getCells(sheet, iRow, posDLen))
        Call col.setPk((Not isBlank(getCells(sheet, iRow, posPK))))
        Call col.setNotNull(getCellsAsBool(sheet, iRow, posNNull))
        Call col.setFk(getCellsAsBool(sheet, iRow, posFk))
        Call col.setDefaultValue(getCells(sheet, iRow, posDefaultVal))
        
        If col.isEmpty() Then Exit Do
        
        If isTableBreakRule(strLogTName, tmpLogTName, strPhyTName, tmpPhyTName) Then
           
            tmpLogTName = strLogTName
            tmpPhyTName = strPhyTName
               
            m_TableId = m_TableId + 1
            ReDim Preserve m_Tables(m_TableId)
            Set m_Tables(m_TableId) = New ErTable
            
            Call m_Tables(m_TableId).setLogicName(strLogTName)
            Call m_Tables(m_TableId).setPhisicName(strPhyTName)
            
            loadTableData = m_TableId
        End If
        If m_TableId >= 0 Then
        
            Dim tmpStr() As String
            tmpStr = separateColName(getCells(sheet, iRow, posDepend))
            tmpDepTName = tmpStr(0)
            tmpDepCName = tmpStr(1)
            
            tmpIsDep = False
            If Not isBlank(tmpDepTName) Then
                tmpIsDep = getCellsAsBool(sheet, iRow, posRelType)
                Call m_Tables(m_TableId).addDependenceTable(tmpDepTName, tmpIsDep)
            End If
            
            If (Not col.isEmpty()) Then
                Call col.setRefTable(tmpDepTName)
                
                '外部キーが参照している列が未設定の場合、同じ列名であると解釈する
                If Util.isBlank(tmpDepCName) Then
                    tmpDepCName = col.getPhysicName
                End If
                
                Call col.setRefCol(tmpDepCName)
                Call col.setDependent(tmpIsDep)
                Call m_Tables(m_TableId).addColObj(col)
            End If
            
            tmpObjType = getCells(sheet, iRow, posObjectType)
            If Not isBlank(tmpObjType) Then
                Call m_Tables(m_TableId).setView(isView(tmpObjType))
            End If
        End If
        
        iRow = iRow + 1
        
        DoEvents
        If Constants.getGlobalCancelFlag Then
            Call Util.showDialog(Constants.INFO_CANCELD_BY_USER)
            Call Log.info(Constants.getMessage(Constants.INFO_CANCELD_BY_USER))
            loadTableData = -1
            GoTo finally
        End If
    Loop While (Not col.isEmpty())
    
finally:
    Call Log.info("<<<<< データの読み込みを完了しました")
    
    ' ============
    '   Logの破棄
    ' ============
    Call Log.terminateLog
    
End Function
'
'
'
Public Function reverseTableData(odbc As ErOdbc, catalog As String, schema As String, tables() As String) As Integer
    Dim schemaParam As ODBCSchemaSearchParam
    Dim tableInfo() As ODBCTableInfo
    Dim colInfo()   As ODBCColumnInfo
    Dim pKeys()     As String
    Dim ret         As Long
    Dim col As ErColumun
    Dim i   As Integer
    Dim j   As Integer
    
    ' ============
    '   Logの準備
    ' ============
    Call Log.initialLog(Util.getLogFilename(ThisWorkbook))
    Call Log.info(">>>>> リバースエンジニアリングによるデータの読み込みを開始します")
    
    
    reverseTableData = 0
    ReDim m_Tables(0)
    Call m_Settings.loadSettings(getWorkBook(), False)

    schemaParam.catalog = catalog
    schemaParam.schema = schema
    
    m_TableId = -1
    
    Call resetODBCSchemaSearchParam(schemaParam)
    For i = LBound(tables) To UBound(tables)
        
        schemaParam.table = tables(i)
        
        m_TableId = m_TableId + 1
        ReDim Preserve m_Tables(m_TableId)

        Set m_Tables(m_TableId) = New ErTable
    
        Call m_Tables(m_TableId).setLogicName(tables(i))
        Call m_Tables(m_TableId).setPhisicName(tables(i))
        
        reverseTableData = m_TableId
        
        'テーブル情報
        ret = odbc.getTables(tableInfo, schemaParam)
        If ret = Constants.ODBC_SUCCESS Then
            If Util.compareIgnoreCase(tableInfo(0).tableType, Constants.ODBC_TYPE_VIEW) Then
                Call m_Tables(m_TableId).setView(True)
            End If
        Else
            Call Log.error("[テーブル情報] " & Constants.getMessage(ret))
        End If
        
        '主キー情報
        ret = odbc.getPrimaryKeys(pKeys, schemaParam)
        If ret = Constants.ODBC_SUCCESS Then
            'NOP
        Else
            Call Log.error("[主キー情報] " & Constants.getMessage(ret))
        End If
        
        'カラム情報
        ret = odbc.getColInfo(colInfo, schemaParam)
        If ret = Constants.ODBC_SUCCESS Then
            For j = LBound(colInfo) To UBound(colInfo)
                Set col = New ErColumun
                
                If isContainArray(RTrim$(colInfo(j).columnName), pKeys) Then
                    Call col.setPk(True)
                End If
                
                Call col.setPos(CInt(val(colInfo(j).ordinalPosition)))
                Call col.setLogicName(colInfo(j).columnName)
                Call col.setPhysicName(colInfo(j).columnName)
                Call col.setDataType(colInfo(j).dataType)
                Call col.setDataLength(odbc.makeDataLength( _
                    colInfo(j).characterMaximumLength, _
                    colInfo(j).numericPrecision, _
                    colInfo(j).numericScale))
                
                Call m_Tables(m_TableId).addColObj(col)
            Next
        Else
            Call Log.error("[カラム情報] " & Constants.getMessage(ret))
        End If
        
        DoEvents
        If Constants.getGlobalCancelFlag Then
            Call Util.showDialog(Constants.INFO_CANCELD_BY_USER)
            Call Log.info(Constants.getMessage(Constants.INFO_CANCELD_BY_USER))
            reverseTableData = -1
            GoTo finally
        End If
    Next
    
finally:
    Call Log.info("<<<<< リバースエンジニアリングによるデータの読み込みを完了しました")
    
    ' ============
    '   Logの破棄
    ' ============
    Call Log.terminateLog
    
End Function
'
'
'
Public Function validate(ByRef msg, ByRef erdInfo As ERDInformation) As Boolean
    Dim result As Boolean
    Dim validator As New ErValidator
    
    result = True
    Call Log.info(">>>>> 整合性チェックを開始します")
    
    result = result And validator.validateDuplicateTabAndCol(msg, m_Tables, erdInfo)
    result = result And validator.validateDataType(msg, m_Tables, erdInfo)
    
    Call Log.info("<<<<< 整合性チェック終了 エラー : " & IIf(result, "なし", "あり"))
    
    validate = result
    
    Set validator = Nothing
End Function
'
'
'
Private Function isTableBreakRule(logicalTableName As String, preLogicalTableName As String, _
                                  physicalTableName As String, prePhysicalTableName As String) As Boolean

    
    If Util.isBlank(logicalTableName & physicalTableName) Then
        isTableBreakRule = False
        Exit Function
    End If
    
    If Not Util.isBlank(logicalTableName) And _
            Not Constants.isEqualKey(logicalTableName, preLogicalTableName) Then
        isTableBreakRule = True
        Exit Function
    End If

    If Not Util.isBlank(physicalTableName) And _
            Not Constants.isEqualKey(physicalTableName, prePhysicalTableName) Then
        isTableBreakRule = True
        Exit Function
    End If

    isTableBreakRule = False
    
End Function
'
'
'
Public Sub writeDDL(ByRef fileName As String, ByRef ddlInfo As DDLInformation)
    Dim writer As New ErDDLWriter
    
    If Not isLoadedTables() Then
        Call Util.showDialog(ERR_NO_LOAD_DDL)
        Exit Sub
    End If
    
    ' ============
    '   Logの準備
    ' ============
    Call Log.initialLog(Util.getLogFilename(ThisWorkbook))
    
    Call Log.info(">>>>> DDLの出力を開始します")
    
    Dim text As String
    Dim selectedButton As Long
    Dim erdInfo As ERDInformation
    erdInfo.mode = Physical
    If Not validate(text, erdInfo) Then
        Call frmMsg.showMessageWindow(Constants.TITLE_MSG_FORM_PROBLEM, _
                                      Constants.MSG_PROBLEM_DETECT, _
                                      text, _
                                      vbYesNo, _
                                      selectedButton)
        If Not (selectedButton = vbYes) Then
            Call Util.showDialog(Constants.INFO_CANCELD_BY_USER)
            GoTo finally
        End If
    End If

    'Dim ddlInfo As DDLInformation
    
    'Call Constants.clearDDLInfo(ddlInfo)
    
    'ddlInfo.sepStr = m_Settings.getDDLSeqStr
    'MsgBox m_Settings.getDDLSeqStr
    'ddlInfo.sepStr = ";"
    
    'ddlInfo.commentStr = m_Settings.getDDLCommentStr

    Dim isOutputRelation As Boolean
    isOutputRelation = Util.intToBool(m_Settings.getOutputRelation)
    
    Dim isOutputLogicalName As Boolean
    isOutputLogicalName = Util.intToBool(m_Settings.getDDLOutputLogicalName)
    
    Call writer.wirteDDL(fileName, m_Tables, ddlInfo, isOutputRelation, isOutputLogicalName)

finally:
    Call Log.info("<<<<< DDLの出力を完了しました")
    
    ' ============
    '   Logの破棄
    ' ============
    Call Log.terminateLog

End Sub
'
'
'
Public Sub drawERD(ByRef sheet As Excel.Worksheet, erdInfo As ERDInformation)
    Dim builder As New ERDBuilder
    Dim i As Integer
    Dim j As Integer
    
    If Not isLoadedTables() Then
        Call Util.showDialog(ERR_NO_LOAD_DDL)
        Exit Sub
    End If
    
    ' ============
    '   Logの準備
    ' ============
    Call Log.initialLog(Util.getLogFilename(ThisWorkbook))
    
    Call Log.info(">>>>> ERDの作成を開始します")
    
    Dim msgText As String
    Dim selectedButton As Long
    If Not validate(msgText, erdInfo) Then
        Call frmMsg.showMessageWindow(Constants.TITLE_MSG_FORM_PROBLEM, _
                                      Constants.MSG_PROBLEM_DETECT, _
                                      msgText, _
                                      vbYesNo, _
                                      selectedButton)
        If Not (selectedButton = vbYes) Then
            Call Util.showDialog(Constants.INFO_CANCELD_BY_USER)
            GoTo finally
        End If
    End If
        
    Call builder.init(sheet)
    Call builder.setFontSize(erdInfo.fontSize)
    
    Dim top As Single
    Dim left As Single
    Dim buttom As Single
    Dim preTop As Single
    
    Dim marginTop   As Integer
    Dim marginLeft  As Integer
    Dim widthlimit  As Integer
    Dim interval    As Integer
    
    marginTop = m_Settings.getMarginTop
    marginLeft = m_Settings.getMarginLeft
    widthlimit = m_Settings.getWidthlimit
    interval = m_Settings.getInterval
    
    top = marginTop
    left = marginLeft
    
    ' 主キー、外部キーを省略する場合
    If Util.intToBool(m_Settings.getOutputElision) Then
        Dim newTables()  As ErTable
        ReDim newTables(UBound(m_Tables))
        
        For i = LBound(m_Tables) To UBound(m_Tables)
            If m_Tables(i).getPKCount() > 0 Or m_Tables(i).getDependTableCount() > 0 Then
                '主キー、関連を持つ場合
                Dim tblObj As ErTable
                Dim colObj As ErColumun
                Dim elised As Boolean: elised = False
                
                
                ' テーブル情報のコピー
                Set tblObj = New ErTable
                Call tblObj.setPhisicName(m_Tables(i).getPhisicName)
                Call tblObj.setLogicName(m_Tables(i).getLogicName)
                Call tblObj.setView(m_Tables(i).isView)
                Call tblObj.setDependent(m_Tables(i).isDependent)
                
                ' 依存情報のコピー
                For j = 0 To (m_Tables(i).getDependTableCount() - 1)
                    Dim depentTable As String
                    Dim dependRelation As Boolean
                    
                    depentTable = m_Tables(i).getDependenceTable(j)
                    dependRelation = m_Tables(i).getDependRelation(j)
                    Call tblObj.addDependenceTable(depentTable, dependRelation)
                Next
                
                ' 列情報のコピー
                For j = 0 To m_Tables(i).getColCount()
                    Set colObj = m_Tables(i).getColumn(j)
                    If colObj.isPk Or colObj.isFk Then
                        Call tblObj.addColObj(colObj)
                    Else
                        elised = True
                    End If
                Next
                
                ' 列を省略した場合に省略記号用の列を追加する。
                If elised Then
                    Dim dummyCol As ErColumun
                    
                    Set dummyCol = New ErColumun
                    Call dummyCol.setData(9999, "...", "...", "", "")
                    Call tblObj.addColObj(dummyCol)
                End If
                
                Set newTables(i) = tblObj
            Else
                ' 主キー、関連を持たない場合
                Set newTables(i) = m_Tables(i)
            End If
        Next
        
        m_Tables = newTables
        
    End If

    
    ' テーブル単位の図の作成
    For i = LBound(m_Tables) To UBound(m_Tables)
        preTop = top
        Call builder.createTable(erdInfo.mode, left, top, m_Tables(i))
        
        If buttom < top Then
            buttom = top
        End If
        
        left = left + interval
        If left > widthlimit Then
            left = marginLeft
            top = buttom + interval
            buttom = 0
        Else
            top = preTop
        End If
        
        DoEvents
        If Constants.getGlobalCancelFlag Then
            Call Util.showDialog(Constants.INFO_CANCELD_BY_USER)
            Call Log.info(Constants.getMessage(Constants.INFO_CANCELD_BY_USER))
            GoTo finally
        End If
    Next
    
    ' リレーションの作成
    If Util.intToBool(m_Settings.getOutputRelation) Then
        Dim parent      As String
        Dim isDepend    As Boolean
        
        For i = LBound(m_Tables) To UBound(m_Tables)
            For j = 0 To (m_Tables(i).getDependTableCount() - 1)
                parent = m_Tables(i).getDependenceTable(j)
                isDepend = m_Tables(i).getDependRelation(j)
                If Not Util.isBlank(parent) Then
                    Call builder.createRelation(parent, m_Tables(i).getPhisicName(), isDepend)
                End If
            Next j
        Next
    End If
    
finally:

    Call Log.info("<<<<< ERDの作成を終了しました")
    
    ' ============
    '   Logの破棄
    ' ============
    Call Log.terminateLog
    
End Sub
'
'
'
Private Function isBlank(str As String) As Boolean
    isBlank = ((Trim$(str)) = "")
End Function
'
'
'
Private Function isView(str As String) As Boolean
    Dim result As Boolean
    
    result = False
    If InStr(1, UCase$(str), "VIEW", vbTextCompare) > 0 Then
        result = True
    End If
    
    isView = result
End Function
'
'
'
Private Function getCells(ByRef sheet As Excel.Worksheet, iRow As Integer, iCol As Integer) As String
    Dim result As String
    
    result = ""
    
    If (iRow > 0 And iCol > 0) Then
        result = CStr(sheet.Cells(iRow, iCol))
    End If
    
    getCells = result
End Function
'
'
'
Private Function getCellsAsBool(ByRef sheet As Excel.Worksheet, iRow As Integer, iCol As Integer) As Boolean
    Dim result As Boolean
    Dim tmp As String
    
    result = False
    tmp = UCase$(getCells(sheet, iRow, iCol))
    
    If (tmp = "TRUE") Or (tmp = "Y") Or (tmp = "YES") Or (tmp = "○") Then
        result = True
    End If
    
    getCellsAsBool = result

End Function
'
'
'
Public Function isLoadedTables() As Boolean
    
    isLoadedTables = (m_TableId >= 0)

End Function
'
' 「テーブル名.列名」を 分解する
'
Public Function separateColName(colname As String) As String()
    Dim names() As String
    Dim result(0 To 1) As String
    Dim i As Integer
    Dim limit As Integer
    
    result(0) = ""
    result(1) = ""
    limit = 1
    
    names = Split(colname, ".")
    If (UBound(names) < limit) Then
        limit = UBound(names)
    End If
        
    For i = 0 To limit
        result(i) = names(i)
    Next

    separateColName = result

End Function
'
'
'
Private Function getCellsAsInt(ByRef sheet As Excel.Worksheet, iRow As Integer, iCol As Integer) As Integer
    getCellsAsInt = CInt(val(getCells(sheet, iRow, iCol)))
End Function
Public Sub setWorkBook(ByRef book As Excel.Workbook)
    Set m_Book = book
End Sub
Private Function getWorkBook() As Excel.Workbook
    Set getWorkBook = m_Book
End Function

Private Sub Class_Initialize()
    m_TableId = -1
    ReDim m_Tables(0)
    Set m_Settings = New Settings
    Set m_Book = ThisWorkbook
End Sub

