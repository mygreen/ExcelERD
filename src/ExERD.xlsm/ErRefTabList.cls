VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErRefTabList"
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
' 外部キーが参照するテーブル情報を管理する
'
'
Private m_TableName     As String
Private m_TableCount    As Integer
Private m_Tables()      As ErTable
Private m_TableNames()  As String
Private m_hasTable      As Boolean
'
'
'
Public Function getTableCount() As Integer
    getTableCount = m_TableCount
End Function
'
'
'
Private Function increaseTableCount() As Integer
    m_TableCount = m_TableCount + 1
    increaseTableCount = getTableCount()
End Function

'
'
'
Public Sub ADD(tableName As String, columnName As String, refColumnName As String)
    Dim table  As ErTable
    Dim colPos As Integer
    
    Call addTable(tableName)
    
    Set table = getTableByName(tableName)
    colPos = table.getColCount + 1
    
    Dim col As New ErColumun
    
    Call col.setPos(colPos)
    Call col.setPhysicName(columnName)
    Call col.setRefCol(refColumnName)
    
    Call table.addColObj(col)
        
    m_hasTable = True
End Sub
'
'
'
Public Sub addTable(tableName As String)
    
    If Not isContain(tableName) Then
        Dim tabId As Integer
        tabId = getTableCount()
        
        ReDim Preserve m_Tables(tabId)
        ReDim Preserve m_TableNames(tabId)
        
        Dim table As New ErTable
        Call table.setPhisicName(tableName)
        
        Set m_Tables(m_TableCount) = table
        m_TableNames(m_TableCount) = tableName
        
        Call increaseTableCount
        
    End If
End Sub
'
'
'
Public Function getTable(index As Integer) As ErTable
    If (LBound(m_Tables) <= index) And (index <= UBound(m_Tables)) Then
        Set getTable = m_Tables(index)
    Else
        Set getTable = New ErTable
    End If
End Function
'
'
'
Public Function getTableByName(tableName As String) As ErTable
    Dim table As New ErTable
    Dim i As Integer
    
    For i = LBound(m_Tables) To UBound(m_Tables)
        If Util.compareIgnoreCase(Trim$(m_Tables(i).getPhisicName), Trim$(tableName)) Then
            Set table = m_Tables(i)
            Exit For
        End If
    Next

    Set getTableByName = table
End Function

'
'
'
Public Function isContain(tableName As String) As Boolean
    Dim i       As Integer
    Dim result  As Boolean
    
    result = False
    For i = LBound(m_TableNames) To UBound(m_TableNames)
        If Util.compareIgnoreCase(Trim$(m_TableNames(i)), tableName) Then
            result = True
            Exit For
        End If
    Next
    
    isContain = result

End Function
'
'
'
Public Function hasTable() As Boolean
    hasTable = m_hasTable
End Function
'
'
'
Public Sub setTableName(tableName As String)
    m_TableName = tableName
End Sub
'
'
'
Public Function getTableName()
    getTableName = m_TableName
End Function
'
'
'
Public Sub clear()
    m_TableName = ""
    m_TableCount = 0
    m_hasTable = False
    ReDim m_Tables(m_TableCount)
    ReDim m_TableNames(m_TableCount)
End Sub
'
'
'
Private Sub Class_Initialize()
    Call clear
End Sub
