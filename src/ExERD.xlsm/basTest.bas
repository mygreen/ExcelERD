Attribute VB_Name = "basTest"
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
' テスト用モジュール
'
'
Sub testSettings()
    Dim st As New Settings
    st.loadSettings ActiveWorkbook, True
    st.saveSettings ActiveWorkbook
End Sub
Sub testDDLLoad()
    Dim doc As New ExERDDocument
    Call doc.loadTableData(ThisWorkbook.Sheets("SAMPLE DDL"))
    
    Dim erInfo As ERDInformation
    erInfo.mode = Physical
    Call doc.drawERD(Sheets("ERD"), erInfo)
End Sub
Sub testDDLPrint()
    Dim doc As New ExERDDocument
    
    Call doc.loadTableData(ThisWorkbook.Sheets("SAMPLE DDL"))
    Call doc.writeDDL("c:\test.sql")
End Sub
Sub testConn()
    Dim conn As Variant
    Dim rs As Variant
    
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    'ODBC non dsn setting
    'conn.Open "DRIVER={Microsoft ODBC for Oracle};SERVER=libra.vishnu.local;uid=oe;pwd=1192"
    
    'ODBC Use Oracle Driver
    conn.Open "dsn=libra;uid=oe;pwd=1192"
    
    rs.Open "select * from all_tables", conn
    
    Dim tname As Variant
    
    Set tname = rs("TABLE_NAME")
    Do Until rs.EOF
        Debug.Print tname
        rs.MoveNext
    Loop
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
Sub testMetaData()
    Dim conn As Variant
    Dim rs As Variant
    Const adSchemaTables = 20
    
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    Dim dns As String
    Dim uid As String
    Dim pwd As String
    Dim catalog As String
    Dim schema As String
    Dim isOracle As Boolean
    
    isOracle = True
    If isOracle Then
        'Oracle
        catalog = Empty
        dns = "libra"
        uid = "OE"
        schema = "OE"
    Else
        'SQL Server
        catalog = "Northwind"
        dns = "brahma"
        uid = "sa"
        schema = "dbo"
    End If
    pwd = "1192"
    
    
    conn.Open "dsn=" & dns & ";uid=" & uid & ";pwd=" & pwd
    
    Set rs = conn.OpenSchema(adSchemaTables, Array(catalog, schema, Empty, Empty))
    
    Dim cnt As Integer
    Do Until rs.EOF
        Debug.Print _
            "TABLE_CATALOG: " & rs!TABLE_CATALOG & vbCr & _
            "TABLE_SCHEMA: " & rs!TABLE_SCHEMA & vbCr & _
            "TABLE_NAME: " & rs!TABLE_NAME & vbCr & _
            "TABLE_TYPE: " & rs!TABLE_TYPE & vbCr

        cnt = cnt + 1
        rs.MoveNext
    Loop
    
    Debug.Print "TABLE CNT : " & cnt
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
