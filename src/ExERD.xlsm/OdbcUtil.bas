Attribute VB_Name = "OdbcUtil"
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

Public Declare Function SQLAllocEnv Lib "odbc32.dll" (ByRef phEnv As Long) As Integer
Public Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal hEnv As Long) As Integer

Public Declare Function SQLDataSources Lib "odbc32.dll" Alias "SQLDataSourcesA" ( _
    ByVal hEnv As Long, _
    ByVal fDirection As Integer, _
    ByVal szDSN As String, _
    ByVal cbDSNMax As Integer, _
    ByRef pcbDSN As Integer, _
    ByVal szDescription As String, _
    ByVal cbDescriptionMax As Integer, _
    ByRef pcbDescription As Integer _
) As Integer

Public Const SQL_SUCCESS = 0
Public Const SQL_ERROR = 1
Public Const SQL_SUCCESS_WITH_INFO = 1
Public Const SQL_NO_DATA_FOUND = 100

Public Const SQL_FETCH_FIRST = 2
Public Const SQL_FETCH_NEXT = 1

Public Const MAX_DSN_LENGTH = 30
Public Const MAX_DSN_DESC_LENGTH = 300


' ADODB
Public Const adSchemaCatalogs = 1
Public Const adSchemaColumns = 4
Public Const adSchemaSchemata = 17
Public Const adSchemaTables = 20
Public Const adSchemaPrimaryKeys = 28
    
    
Public Type ODBCSearchInfo
    dns     As String
    uid     As String
    pwd     As String
    catalog As Variant
    schema  As Variant
    table   As Variant
End Type


Public Function getODBCDataSourceNames(ByRef dsnList() As String)
    Dim hEnv            As Long
    Dim szDSN           As String * MAX_DSN_LENGTH
    Dim cbDSN           As Integer
    Dim szDescription   As String * MAX_DSN_DESC_LENGTH
    Dim cbDescription   As Integer
    Dim retcode         As Integer
    Dim dsnPos          As Integer
    
On Error GoTo errhandler

    retcode = SQLAllocEnv(hEnv)
    If retcode <> SQL_SUCCESS Then
        getODBCDataSourceNames = False
        Exit Function
    End If
    
    dsnPos = 0
    ReDim dsnList(dsnPos)
    retcode = SQLDataSources(hEnv, SQL_FETCH_FIRST, _
                szDSN, MAX_DSN_LENGTH, cbDSN, _
                szDescription, MAX_DSN_DESC_LENGTH, cbDescription)
            
    Do While (retcode <> SQL_ERROR _
        And retcode <> SQL_NO_DATA_FOUND)
        
        ReDim Preserve dsnList(dsnPos)
        dsnList(dsnPos) = LeftB(szDSN, InStrB(szDSN, vbNullChar))
        dsnPos = dsnPos + 1
    
        retcode = SQLDataSources(hEnv, SQL_FETCH_NEXT, _
                    szDSN, MAX_DSN_LENGTH, cbDSN, _
                    szDescription, MAX_DSN_DESC_LENGTH, cbDescription)
    Loop
    
    retcode = SQLFreeEnv(hEnv)
    getODBCDataSourceNames = IIf(retcode = SQL_SUCCESS, True, False)
    
    Exit Function
errhandler:
    retcode = SQLFreeEnv(hEnv)
    getODBCDataSourceNames = False

End Function

Public Sub initOdbcSearchInfo(odbcInfo As ODBCSearchInfo)
    With odbcInfo
        .dns = Empty
        .uid = Empty
        .pwd = Empty
        .catalog = Empty
        .schema = Empty
        .table = Empty
    End With

End Sub
Public Function getCategories(ByRef categories() As String, _
                                odbcInfo As ODBCSearchInfo) As Boolean
                                
    Dim conn        As Variant
    Dim rs          As Variant
    Dim retPos      As Integer

On Error GoTo errhandler

    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    conn.Open "dsn=" & odbcInfo.dns & ";uid=" & odbcInfo.uid & ";pwd=" & odbcInfo.pwd

    retPos = 0
    ReDim categories(retPos)
    
    Set rs = conn.OpenSchema(adSchemaCatalogs)
    Do Until rs.EOF
        
        ReDim Preserve categories(retPos)
        categories(retPos) = rs!CATALOG_NAME
        retPos = retPos + 1
            
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    getCategories = True
    Exit Function

errhandler:
    
    getCategories = False
End Function
Public Function getTables(ByRef tables() As String, _
                                odbcInfo As ODBCSearchInfo) As Boolean
                                
    Dim conn        As Variant
    Dim rs          As Variant
    Dim retPos      As Integer

On Error GoTo errhandler

    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    conn.Open "dsn=" & odbcInfo.dns & ";uid=" & odbcInfo.uid & ";pwd=" & odbcInfo.pwd

    retPos = 0
    ReDim tables(retPos)
    
    Set rs = conn.OpenSchema(adSchemaTables, Array(odbcInfo.catalog, odbcInfo.schema, odbcInfo.table, Empty))
    Do Until rs.EOF
        
        ReDim Preserve tables(retPos)
        tables(retPos) = rs!TABLE_NAME
        retPos = retPos + 1
            
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    getTables = True
    Exit Function

errhandler:
    
    getTables = False
End Function






'
' ODBC データ ソースの一覧を取得
' @see http://support.microsoft.com/default.aspx?scid=kb;ja;119064
'
'
'   #include <afxcoll.h>    //Needed for CStringList MFC class.
'   #include "odbcinst.h"
'   #include "sql.h"
'   #include "sqlext.h"
'
'   // NOTE: in 16-bit Visual C++ link with odbcinst.lib
'   //       in 32-bit Visual C++ 2.x link with odbccp32.lib
'   //       in 32-bit Visual C++ 4.x no need to change link options
'
'   #define MAX_DSN_LENGTH 30
'   #define MAX_DSN_DESC_LENGTH 300
'
'   BOOL GetODBCDataSourceNames(CStringList * pList)
'   {
'       HENV hEnv;
'       char szDSN[MAX_DSN_LENGTH];
'       SWORD cbDSN;
'       UCHAR szDescription[MAX_DSN_DESC_LENGTH];
'       SWORD cbDescription;
'       RETCODE retcode;
'
'       ASSERT(pList->IsEmpty());
'       if (SQLAllocEnv(&hEnv)!=SQL_SUCCESS)
'           return FALSE;
'
'       while (retcode=SQLDataSources(hEnv, SQL_FETCH_NEXT,
'                    (UCHAR FAR *) &szDSN, MAX_DSN_LENGTH, &cbDSN,
'                    (UCHAR FAR *) &szDescription,MAX_DSN_DESC_LENGTH,
'                     &cbDescription) != SQL_NO_DATA_FOUND
'                    &&retcode!=SQL_ERROR)
'
'          {
'               pList->AddTail(szDSN);
'          }
'
'       SQLFreeEnv(hEnv);
'       if (retcode==SQL_ERROR)
'         return FALSE;
'
'       return TRUE;
'   }

