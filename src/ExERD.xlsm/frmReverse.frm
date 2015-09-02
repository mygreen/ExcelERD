VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReverse 
   Caption         =   "ReverseForm"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4530
   OleObjectBlob   =   "frmReverse.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmReverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
' リバースエンジニアリングER図作成用画面
'
'
Private prop            As Settings                 '設定値保持
Private odbc            As New ErOdbc               'ODBC操作クラス
Private m_schemaParam   As ODBCSchemaSearchParam    'ODBCスキーマ検索情報
Private isRunning   As Boolean

'
' リバースエンジニアリングER図作成用画面表示
'
Public Sub showReverseDialog(Settings As Settings, Optional ByRef dataTypeFileName As String)
    Set prop = Settings
    
    If Not odbc.loadDataTypeFile(dataTypeFileName) Then
        Call Util.showDialog(Constants.INFO_CANCELD_BY_USER)
        Exit Sub
    End If
    
    Call initControl
    Call Me.show(vbModal)
    Unload Me
    
End Sub
'
' コントロールの初期化
'
Private Sub initControl()
    Call loadDsn
End Sub
'
' ODBCデータソースの読み込み
'
Private Function loadDsn() As Boolean
    Dim dnsList()   As String
    Dim result      As Boolean
    Dim i           As Integer
    
    result = OdbcUtil.getODBCDataSourceNames(dnsList)
    If result Then
        For i = 0 To UBound(dnsList)
            Call cmbDsn.AddItem(dnsList(i), i)
        Next
        'If LBound(dnsList) >= 0 Then
        '    cmbDsn.text = dnsList(0)
        'End If
    End If
    loadDsn = result
End Function
'
' キャンセル処理
'
Private Sub cmdCancel_Click()
    If isRunning Then
        If Util.showDialog(Q_YN_CANCEL_PROC) = vbYes Then
            Call Constants.setGlobalCancelFlag(True)
        End If
    Else
        Me.Hide
    End If
End Sub
'
' 作成処理
'
Private Sub cmdCreateErd_Click()
    Dim wrkBook     As Workbook
    Dim sheetInfo   As SheetInformation
    Dim erdInfo     As ERDInformation
    Dim erdDoc      As New ExERDDocument
    Dim erdSheet    As Worksheet
    Dim mode        As ERDMODE
    Dim errResult   As Long
    Dim tables()    As String
    Dim i           As Integer
    Dim tblCnt      As Integer
    Dim hasAnyTable As Boolean
    
    Call lockControls(True)
    
    tblCnt = 0
    hasAnyTable = False
    For i = 0 To lstTables.ListCount - 1
        If lstTables.selected(i) Then
            hasAnyTable = True
            ReDim Preserve tables(tblCnt)
            tables(tblCnt) = removeTypePrefix(lstTables.List(i))
            tblCnt = tblCnt + 1
        End If
    Next
    If Not hasAnyTable Then
        Util.showDialog (Constants.ERR_NO_TABLE)
        multiPageReverse.Value = 1
        GoTo finally
    End If
    
    ' ---------- ERD Sheet Select ----------
    Call Constants.clearSheetInfo(sheetInfo)
    Call frmSelSheetDlg.showERDSheetSelectDialog(sheetInfo)
    errResult = validateSelectedSheet(sheetInfo)
    If errResult <> Util.NO_ERROR Then
        Call Util.showDialog(errResult)
        GoTo finally
    End If
    
    If sheetInfo.isNewSheet Then
        Set wrkBook = Application.Workbooks(sheetInfo.bookName)
        Set erdSheet = wrkBook.Sheets.ADD
    Else
        Set erdSheet = Application.Workbooks( _
                        sheetInfo.bookName).Worksheets(sheetInfo.sheetName)
    End If
    
    '---------- Create ERD ----------
    If erdDoc.reverseTableData(odbc, _
                                 cmbCatalogs.text, _
                                 cmbSchemas.text, _
                                 tables) <> -1 Then
        erdInfo.mode = Physical
        erdInfo.fontSize = prop.getFontSize
        Call erdDoc.drawERD(erdSheet, erdInfo)
    End If
    
    Call MsgBox("終了しました。", vbInformation, Constants.getAppInfo)

finally:
    Call lockControls(False)
    
End Sub
'
'
'
Private Sub lockControls(isLock As Boolean)
    isRunning = isLock
    Call Constants.setGlobalCancelFlag(False)
    multiPageReverse.Enabled = Not isLock
    cmdCreateErd.Enabled = Not isLock
End Sub
'
' 接続処理
'
Private Sub cmdConnect_Click()
    Dim dsn As String
    Dim ret As Long
    
    dsn = cmbDsn.text
    ret = odbc.openConnection(dsn, txtUid.text, txtPwd.text)
    If ret = Constants.ODBC_SUCCESS Then
        cmbCatalogs.clear
        cmbSchemas.clear
        lstTables.clear
        
        cmbCatalogs.Enabled = True
        cmbSchemas.Enabled = True
        cmdOpenCatalog.Enabled = True
        cmdOpenSchema.Enabled = True
        
        ret = loadCatalogs
        If ret <> Constants.ODBC_SUCCESS Then
            cmbCatalogs.Enabled = False
            cmdOpenCatalog.Enabled = False
            
            If ret = Constants.ERR_ODBC_NOT_SUPPORTED_OPERATION Then
                Call loadSchemas(cmbCatalogs.text)
            Else
                Util.showDialog (ret)
            End If
            
        End If
        multiPageReverse.Value = 1
    Else
        Call Util.showDialog(ret)
    End If
    
End Sub
'
' カタログ選択処理
'
Private Sub cmdOpenCatalog_Click()
    Dim ret As Long
    
    ret = loadSchemas(cmbCatalogs.text)
    If ret <> Constants.ODBC_SUCCESS Then
        cmbSchemas.clear
        cmbSchemas.Enabled = False
        cmdOpenSchema.Enabled = False

        If ret = Constants.ERR_ODBC_NOT_SUPPORTED_OPERATION Then
            Call loadTables(cmbCatalogs.text, cmbSchemas.text)
        Else
            Util.showDialog (ret)
        End If
    End If

End Sub
'
' スキーマ選択処理
'
Private Sub cmdOpenSchema_Click()
    Dim ret As Long
    
    ret = loadTables(cmbCatalogs.text, cmbSchemas.text)
    If ret <> Constants.ODBC_SUCCESS Then
        cmbSchemas.clear
        If ret = Constants.ERR_ODBC_NOT_SUPPORTED_OPERATION Then
            'NOP
        Else
            Util.showDialog (ret)
        End If
    End If

End Sub
'
' カタログ情報の読み込み
'
Private Function loadCatalogs() As Long
    Dim ret         As Long
    Dim catalogs()  As String
    Dim i           As Integer
    
    Call cmbCatalogs.clear
    ret = odbc.getCatalogs(catalogs, m_schemaParam)
    If ret = Constants.ODBC_SUCCESS Then
        For i = 0 To UBound(catalogs)
            Call cmbCatalogs.AddItem(catalogs(i))
        Next
    End If
    
    loadCatalogs = ret
    
End Function
'
' スキーマ情報の読み込み
'
Private Function loadSchemas(catalog As String) As Long
    Dim ret         As Long
    Dim schemas()   As String
    Dim i           As Integer
    
    Call cmbSchemas.clear
    m_schemaParam.catalog = catalog
    ret = odbc.getSchemas(schemas, m_schemaParam)
    If ret = Constants.ODBC_SUCCESS Then
        For i = 0 To UBound(schemas)
            Call cmbSchemas.AddItem(schemas(i))
        Next
    End If
    
    loadSchemas = ret
    
End Function
'
' テーブル情報の読み込み
'
Private Function loadTables(catalog As String, schema As String) As Long
    Dim ret         As Long
    Dim tableInfo() As ODBCTableInfo
    Dim i           As Integer
    
    Call lstTables.clear
    m_schemaParam.catalog = catalog
    m_schemaParam.schema = schema
    ret = odbc.getTables(tableInfo, m_schemaParam)
    If ret = Constants.ODBC_SUCCESS Then
        For i = 0 To UBound(tableInfo)
            Call lstTables.AddItem(encloseType(tableInfo(i).tableType) & tableInfo(i).tableName, i)
        Next
    End If
    
    loadTables = ret
    
End Function
'
'
'
Private Function removeTypePrefix(tableName As String) As String
    
    If InStr(1, tableName, encloseType(Constants.ODBC_TYPE_TABLE), vbTextCompare) = 1 Then
        removeTypePrefix = Mid$(tableName, Len(encloseType(Constants.ODBC_TYPE_TABLE)) + 1)
        Exit Function
    End If

    If InStr(1, tableName, encloseType(Constants.ODBC_TYPE_VIEW), vbTextCompare) = 1 Then
        removeTypePrefix = Mid$(tableName, Len(encloseType(Constants.ODBC_TYPE_VIEW)) + 1)
        Exit Function
    End If

End Function
Private Function encloseType(tableType As String) As String
    
    encloseType = "[" & tableType & "] "

End Function
'
'
'
Private Sub UserForm_Initialize()
    Me.Caption = Constants.TITLE_REVERSE_FORM
    Call clearODBCSchemaSearchParam(m_schemaParam)
End Sub

Private Sub UserForm_Terminate()
    odbc.closeConnection
End Sub
