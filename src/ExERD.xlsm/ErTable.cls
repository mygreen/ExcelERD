VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErTable"
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
' RDBにおけるテーブルを表現するクラス
'
'
'

Private m_physicName        As String
Private m_logicName         As String
Private m_IsView            As Boolean
Private m_colCount          As Integer
Private m_dependTableCount  As Integer
Private m_IsDependent       As Boolean

Private colmns() As ErColumun
Private dependTables() As String
Private dependRelations() As Boolean


'
' カラム数を返す
'
Public Function getColCount() As Integer
    getColCount = m_colCount
End Function
'
' カラムの数をインクリメント
'
Private Function increaseColCount() As Integer
    m_colCount = m_colCount + 1
    increaseColCount = getColCount()
End Function
'
'
'
Public Function getDependTableCount() As Integer
    getDependTableCount = m_dependTableCount
End Function
'
'
'
Private Function increaseDependTableCount() As Integer
    m_dependTableCount = m_dependTableCount + 1
    increaseDependTableCount = getDependTableCount()
End Function

'
' テーブルクラスの属性を一括設定
'
Public Sub setData(logicName As String, phisicName As String, Optional isDependent As Boolean, Optional isView As Boolean)
    Call setLogicName(logicName)
    Call setPhisicName(phisicName)
    Call setDependent(isDependent)
    Call setView(isView)
    
End Sub
'
' テーブルクラスにカラムを追加
'
Public Sub addColumn(pos As Integer, logicName As String, _
                     phisicName As String, _
                     dataType As String, _
                     dataLength As String, _
                     Optional isPk As Boolean, _
                     Optional isFk As Boolean, _
                     Optional isDependent As Boolean)
    Dim colId As Integer
    colId = getColCount
    
    ReDim Preserve colmns(colId)
    Set colmns(colId) = New ErColumun
    Call colmns(colId).setData(pos, logicName, phisicName, dataType, dataLength, isPk, isFk)

    If isDependent Then
        Call setDependent(True)
    End If
    
    Call increaseColCount
End Sub
'
' テーブルクラスにカラムを追加
'
Public Sub addColObj(ByVal colObj As ErColumun)
    Dim colId As Integer
    colId = getColCount
    
    ReDim Preserve colmns(colId)
    Set colmns(colId) = colObj
    
    If colObj.isDependent Then
        Call setDependent(True)
    End If
    
    Call increaseColCount
    
End Sub
'
' テーブルクラスが保持するカラムの名称最大長を取得
'
Public Function getMaxColNameLength(mode As ERDMODE) As Integer
    Dim result As Integer
    Dim i As Integer
    
    result = 0
    
    For i = LBound(colmns) To UBound(colmns)
        If colmns(i).getColNameLength(mode) > result Then
            result = colmns(i).getColNameLength(mode)
        End If
    Next
    
    getMaxColNameLength = result
End Function
'
' テーブルクラスが保持するデータ型の名称最大長を取得
'
Public Function getMaxDataTypeLength() As Integer
    Dim result As Integer
    Dim i As Integer
    
    result = 0
    
    For i = LBound(colmns) To UBound(colmns)
        If colmns(i).getDataTypeLen() > result Then
            result = colmns(i).getDataTypeLen()
        End If
    Next
    
    getMaxDataTypeLength = result
End Function
'
' カラムの文字数を取得する(名称＋データ型)
'
Public Function getColStringLength(mode As ERDMODE) As Integer
    getColStringLength = getMaxColNameLength(mode) + Constants.SEP_MARGIN + getMaxDataTypeLength()
End Function
'
'
'
Public Function getTableName(mode As ERDMODE) As String
    Dim result As String
    result = ""
    If mode = ERDMODE.Physical Then
        result = getPhisicName()
    ElseIf mode = ERDMODE.Logical Then
        result = getLogicName()
    ElseIf mode = ERDMODE.PhysicalAndLogical Then
        result = getPhisicName() & " : " & getLogicName()
    Else
        result = getPhisicName()
    End If
    getTableName = result
End Function
'
'
'
Public Function getPKCount() As Integer
    Dim result As Integer
    Dim i As Integer
    
    result = 0
    For i = LBound(colmns) To UBound(colmns)
        If colmns(i).isPk Then
            result = result + 1
        End If
    Next

    getPKCount = result
End Function
'
'
'
Public Function getSerialColumnString(mode As ERDMODE) As String
    Dim i As Integer
    Dim result As String
    Dim colNameLen As Integer
    Dim colDataLen As Integer
    Dim pkCnt      As Integer
    
    Call sortColumn
    
    colNameLen = getMaxColNameLength(mode)
    colDataLen = getMaxDataTypeLength()
    pkCnt = getPKCount()
    result = ""
    For i = LBound(colmns) To UBound(colmns)
        result = result & colmns(i).toColmunString(mode, colNameLen) & Chr(10)
        If (i + 1) = pkCnt Then
            result = result & Chr(10)
        End If
    Next

    getSerialColumnString = result
End Function
Public Function getColumnString(mode As ERDMODE, index As Integer) As String
    Dim result As String
    result = ""
    
    If (LBound(colmns) <= index) And (index <= UBound(colmns)) Then
        result = colmns(index).toColmunString(mode, getMaxColNameLength(mode))
    End If

    getColumnString = result
End Function
Public Function getColumnName(mode As ERDMODE, index As Integer) As String
    Dim result As String
    result = ""
    
    If (LBound(colmns) <= index) And (index <= UBound(colmns)) Then
        result = colmns(index).getColName(mode)
    End If

    getColumnName = result
End Function
'
'
'
Public Function getColumn(index As Integer) As ErColumun
    
    If (LBound(colmns) <= index) And (index <= UBound(colmns)) Then
        Set getColumn = colmns(index)
    Else
        Set getColumn = New ErColumun
    End If

End Function
'
'
Public Sub addDependenceTable(tableName As String, isDependRelation As Boolean)
    Dim isHit As Boolean
    Dim i As Integer
    
    isHit = False
    For i = LBound(dependTables) To UBound(dependTables)
        If Constants.keyRule(tableName) = dependTables(i) Then
            isHit = True
            Exit For
        End If
    Next
    
    If Not isHit Then
        Dim dependsId As Integer
        dependsId = getDependTableCount()
        
        ReDim Preserve dependTables(dependsId)
        ReDim Preserve dependRelations(dependsId)
        
        dependTables(dependsId) = Constants.keyRule(tableName)
        dependRelations(dependsId) = isDependRelation
        
        Call increaseDependTableCount
    End If
    
End Sub
'
'
'
Public Function getDependenceTable(index As Integer)
    Dim result As String
    result = ""
    
    If (LBound(dependTables) <= index) And (index <= UBound(dependTables)) Then
        result = dependTables(index)
    End If

    getDependenceTable = result
End Function
'
'
'
Public Function getDependRelation(index As Integer) As Boolean
    Dim result As Boolean
    result = False
    
    If (LBound(dependTables) <= index) And (index <= UBound(dependTables)) Then
        result = dependRelations(index)
    End If

    getDependRelation = result
End Function
'
' カラムを PK + POS 順にソートする
'
Public Sub sortColumn()
    Dim tmp As ErColumun
    Dim i As Integer
    Dim j As Integer
    
    For i = LBound(colmns) To UBound(colmns)
        For j = i + 1 To UBound(colmns)
            If StrComp(colmns(i).getSortKey(), colmns(j).getSortKey(), vbTextCompare) > 0 Then
                Set tmp = colmns(i)
                Set colmns(i) = colmns(j)
                Set colmns(j) = tmp
            End If
        Next
    Next

End Sub
' ===== アクセサ ====
Public Sub setPhisicName(newVal As String)
    m_physicName = newVal
End Sub
Public Function getPhisicName() As String
    getPhisicName = m_physicName
End Function
Public Sub setLogicName(newVal As String)
    m_logicName = newVal
End Sub
Public Function getLogicName() As String
    getLogicName = m_logicName
End Function
Public Sub setView(isView As Boolean)
    m_IsView = isView
End Sub
Public Function isView() As Boolean
    isView = m_IsView
End Function
Public Sub setDependent(isDependent As Boolean)
    m_IsDependent = isDependent
End Sub
Public Function isDependent() As Boolean
    isDependent = m_IsDependent
End Function

'
'
'
Private Sub Class_Initialize()
    m_physicName = ""
    m_logicName = ""
    m_colCount = 0
    m_dependTableCount = 0
    m_IsView = False
    m_IsDependent = False
    ReDim colmns(m_colCount)
    ReDim dependTables(m_dependTableCount)
    
End Sub

