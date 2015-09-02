VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErOdbc"
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
' ADOを利用しODBCのメタ情報を利用するクラス
'
'

Private conn                As Variant  'Connectionオブジェクト
Private isOpen              As Boolean  'Connectionが開かれているか
Private dataTypes(255)      As String   'データ型文字列
Private m_dataTypeFileName  As String   'データ型マッピングファイル名
'
' ODBCデータソースに接続し、Connectionを開く
'
Public Function openConnection(dsn As String, uid As String, pwd As String) As Long
    On Error GoTo errhandler
    
    Call closeConnection
    
    Set conn = CreateObject(Constants.ODBC_ADO_CONN_STR)
    conn.Open "dsn=" & dsn & ";uid=" & uid & ";pwd=" & pwd

    isOpen = True
    openConnection = Constants.ODBC_SUCCESS
    
    Exit Function

errhandler:
    If Err.Number < 0 Then
        Call MsgBox(Err.Number & vbCrLf & Err.Description, vbCritical, Constants.getAppInfo)
    End If
    openConnection = handleOdbcError(Err.Number, Constants.ERR_ODBC_CONNECT_FAIL)
    
End Function
'
' Connectionを閉じる
'
Public Function closeConnection() As Long
    On Error GoTo errhandler
    
    isOpen = False
    
    If Not isEmpty(conn) Then
        Call conn.Close
    End If
    
    closeConnection = Constants.ODBC_SUCCESS
    
    Exit Function
errhandler:
    If Err.Number < 0 Then
        Call MsgBox(Err.Number & vbCrLf & Err.Description, vbCritical, Constants.getAppInfo)
    End If
    closeConnection = handleOdbcError(Err.Number)
    
End Function
'
' カタログ情報を取得する
'
Public Function getCatalogs(ByRef catalogs() As String, _
                                odbcInfo As ODBCSchemaSearchParam) As Long
                                
    If Not isOpen Then
        getCatalogs = Constants.ERR_ODBC_NO_CONNECTION
        Exit Function
    End If
    
    Dim rs          As Variant
    Dim retPos      As Integer

On Error GoTo errhandler
    retPos = 0
    ReDim catalogs(retPos)
    
    Call resetODBCSchemaSearchParam(odbcInfo)
    Set rs = CreateObject(Constants.ODBC_ADO_RECORDSET)
    
    Set rs = conn.OpenSchema(adSchemaCatalogs)
    Do Until rs.EOF
        ReDim Preserve catalogs(retPos)
        catalogs(retPos) = rs!CATALOG_NAME
        retPos = retPos + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    
    getCatalogs = Constants.ODBC_SUCCESS
    
    Exit Function
errhandler:
    If Err.Number < 0 Then
        Call MsgBox(Err.Number & vbCrLf & Err.Description, vbCritical, Constants.getAppInfo)
    End If
    getCatalogs = handleOdbcError(Err.Number)
    
End Function
'
' スキーマ情報を取得する
'
Public Function getSchemas(ByRef schemas() As String, _
                                odbcInfo As ODBCSchemaSearchParam) As Long
                                
    If Not isOpen Then
        getSchemas = Constants.ERR_ODBC_NO_CONNECTION
        Exit Function
    End If
    
    Dim rs          As Variant
    Dim retPos      As Integer

On Error GoTo errhandler
    retPos = 0
    ReDim schemas(retPos)
    
    Call resetODBCSchemaSearchParam(odbcInfo)
    Set rs = CreateObject(Constants.ODBC_ADO_RECORDSET)
    
    Set rs = conn.OpenSchema(adSchemaSchemata, Array(odbcInfo.catalog, Empty, Empty))
    Do Until rs.EOF
        ReDim Preserve schemas(retPos)
        schemas(retPos) = rs!SCHEMA_NAME
        retPos = retPos + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    
    getSchemas = Constants.ODBC_SUCCESS
    
    Exit Function
errhandler:
    If Err.Number < 0 Then
        Call MsgBox(Err.Number & vbCrLf & Err.Description, vbCritical, Constants.getAppInfo)
    End If

    getSchemas = handleOdbcError(Err.Number)
    
End Function
'
' テーブル情報を取得する
'
Public Function getTables(ByRef tableInfo() As ODBCTableInfo, _
                                odbcInfo As ODBCSchemaSearchParam) As Long
                                
    If Not isOpen Then
        getTables = Constants.ERR_ODBC_NO_CONNECTION
        Exit Function
    End If
    
    Dim rs          As Variant
    Dim retPos      As Integer
    Dim tableType   As String
    
On Error GoTo errhandler
    retPos = 0
    ReDim tableInfo(retPos)

    Call resetODBCSchemaSearchParam(odbcInfo)
    Set rs = CreateObject(Constants.ODBC_ADO_RECORDSET)
    
    Set rs = conn.OpenSchema(adSchemaTables, Array(odbcInfo.catalog, odbcInfo.schema, odbcInfo.table, Empty))
    Do Until rs.EOF
        tableType = rs!TABLE_TYPE
        
        If Util.compareIgnoreCase(tableType, Constants.ODBC_TYPE_TABLE) _
            Or Util.compareIgnoreCase(tableType, Constants.ODBC_TYPE_VIEW) Then
    
            ReDim Preserve tableInfo(retPos)
            tableInfo(retPos).tableName = rs!TABLE_NAME
            tableInfo(retPos).tableType = tableType
            
            retPos = retPos + 1
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    
    getTables = Constants.ODBC_SUCCESS
    
    Exit Function
errhandler:
    If Err.Number < 0 Then
        Call MsgBox(Err.Number & vbCrLf & Err.Description, vbCritical, Constants.getAppInfo)
    End If

    getTables = handleOdbcError(Err.Number)
    
End Function
'
' 主キー情報を取得する
'
Public Function getPrimaryKeys(ByRef pKeys() As String, _
                                odbcInfo As ODBCSchemaSearchParam) As Long
                                
    If Not isOpen Then
        getPrimaryKeys = Constants.ERR_ODBC_NO_CONNECTION
        Exit Function
    End If
    
    Dim rs          As Variant
    Dim retPos      As Integer

On Error GoTo errhandler
    retPos = 0
    ReDim pKeys(retPos)

    Call resetODBCSchemaSearchParam(odbcInfo)
    Set rs = CreateObject(Constants.ODBC_ADO_RECORDSET)
    
    Set rs = conn.OpenSchema(adSchemaPrimaryKeys, Array(odbcInfo.catalog, odbcInfo.schema, odbcInfo.table))
    Do Until rs.EOF
        ReDim Preserve pKeys(retPos)
        pKeys(retPos) = rs!COLUMN_NAME
        
        retPos = retPos + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    
    getPrimaryKeys = Constants.ODBC_SUCCESS
    
    Exit Function
errhandler:
    If Err.Number < 0 Then
        Call MsgBox(Err.Number & vbCrLf & Err.Description, vbCritical, Constants.getAppInfo)
    End If

    getPrimaryKeys = handleOdbcError(Err.Number)
    
End Function
'
' カラム情報を取得する
'
Public Function getColInfo(ByRef colInfo() As ODBCColumnInfo, _
                            odbcInfo As ODBCSchemaSearchParam) As Long
                                
    If Not isOpen Then
        getColInfo = Constants.ERR_ODBC_NO_CONNECTION
        Exit Function
    End If
    
    Dim rs          As Variant
    Dim retPos      As Integer

On Error GoTo errhandler
    retPos = 0
    ReDim colInfo(retPos)

    Call resetODBCSchemaSearchParam(odbcInfo)
    Set rs = CreateObject(Constants.ODBC_ADO_RECORDSET)
    
    Set rs = conn.OpenSchema(adSchemaColumns, Array(odbcInfo.catalog, odbcInfo.schema, odbcInfo.table, Empty))
    Do Until rs.EOF
        ReDim Preserve colInfo(retPos)
        With colInfo(retPos)
            .ordinalPosition = rs!ORDINAL_POSITION.Value
            .columnName = rs!COLUMN_NAME.Value
            .dataType = getDataTypeName(rs!DATA_TYPE.Value)
            .characterMaximumLength = IIf(IsNull(rs!CHARACTER_MAXIMUM_LENGTH), "", rs!CHARACTER_MAXIMUM_LENGTH.Value)
            .numericPrecision = IIf(IsNull(rs!NUMERIC_PRECISION), "", rs!NUMERIC_PRECISION.Value)
            .numericScale = IIf(IsNull(rs!NUMERIC_SCALE), "", rs!NUMERIC_SCALE.Value)
        End With
        retPos = retPos + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    
    getColInfo = Constants.ODBC_SUCCESS
    
    Exit Function
errhandler:
    If Err.Number < 0 Then
        Call MsgBox(Err.Number & vbCrLf & Err.Description, vbCritical, Constants.getAppInfo)
    End If
    
    getColInfo = handleOdbcError(Err.Number)
                                
End Function
'
' カラム情報のデータ長から、データ長文字列を生成する
'
Public Function makeDataLength(characterMaximumLength As String, _
                               numericPrecision As String, _
                               numericScale As String) As String
    Dim result As String
    result = ""
    If Not Util.isBlank(characterMaximumLength) Then
        result = characterMaximumLength
    Else
        If Not Util.isBlank(numericPrecision) Then
            result = numericPrecision
            If Not Util.isBlank(numericScale) Then
                result = result & "," & numericScale
            End If
        End If
    End If

    makeDataLength = result
End Function
'
' ODBCデータ型定数を設定ファイルのデータ型名に変換する
'
Public Function getDataTypeName(dataIndex As Integer) As String
    
    If 0 <= dataIndex And dataIndex <= 255 Then
        getDataTypeName = dataTypes(dataIndex)
        Exit Function
    End If
    getDataTypeName = ""
    
End Function
'
' ODBCデータ型定数マッピングファイルを読み込む
'
Public Function loadDataTypeFile(ByRef dataTypeFileName As String) As Boolean
    Dim openFileName    As Variant
    Dim i               As Integer
    
    If isBlank(dataTypeFileName) Or Dir$(dataTypeFileName) = "" Then
        Call Util.showDialog(Constants.ERR_NO_DATATYPEFILE)
        
        openFileName = Util.chooseFile(REVERSE_DATATYPEFILE_FILTER, ThisWorkbook.Path)
        If openFileName <> False Then
            dataTypeFileName = openFileName
        Else
            loadDataTypeFile = False
            Exit Function
        End If
    End If
    
    For i = 0 To 255
        dataTypes(i) = Util.getProperty(CStr(i), CStr(i), dataTypeFileName)
    Next
        
    loadDataTypeFile = True
End Function
'
' ADOで発生するエラーのうち必要なものをハンドルする
'
Public Function handleOdbcError(errNumber As Long, Optional defaultNumber As Long) As Long
    Dim result As Long
    
    Select Case errNumber
    Case 429    'ActiveXコンポーネントはオブジェクトを作成できません
        result = ERR_ODBC_ADO_LOADING_FAIL
    Case 3251   'オブジェクトまたはプロバイダは要求された操作を実行できません
        result = Constants.ERR_ODBC_NOT_SUPPORTED_OPERATION
    Case Else
        If IsMissing(defaultNumber) Then
            result = Constants.ERR_ODBC_GENERAL
        Else
            result = defaultNumber
        End If
    End Select

    handleOdbcError = result
End Function
'
'
'
Private Sub Class_Terminate()
    Set conn = Nothing
End Sub
'
' ODBCデータ型定数マッピングファイルを取得
'
Public Function getDataTypeFileName() As String
    getDataTypeFileName = m_dataTypeFileName
End Function
'
' ODBCデータ型定数マッピングファイルを設定
'
Public Sub setDataTypeFileName(dataTypeFileName As String)
    m_dataTypeFileName = dataTypeFileName
End Sub
