VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErValidator"
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
' 表名、列名の重複チェック
'
Public Function validateDuplicateTabAndCol(ByRef msg, ByRef m_Tables() As ErTable, ByRef erdInfo As ERDInformation) As Boolean
    Dim result As Boolean
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim resultTab        As String
    Dim resultCol        As String
    
    Dim colCnt           As Integer
    Dim colmuns()        As String
    Dim tables()         As String
    Dim duplicateItems() As String
    Dim strTmp           As String
    
    result = True
    ReDim tables(UBound(m_Tables))
    For i = LBound(m_Tables) To UBound(m_Tables)
        tables(i) = m_Tables(i).getTableName(erdInfo.mode)
    
        colCnt = m_Tables(i).getColCount
        ReDim colmuns(colCnt)
        
        For j = 1 To colCnt
            colmuns(j) = m_Tables(i).getColumn(j - 1).getColName(erdInfo.mode)
        Next
        
        If Util.duplicatedCheck(colmuns, duplicateItems) Then
            result = False
            For k = LBound(duplicateItems) To UBound(duplicateItems)
                strTmp = "[列名の重複] " & m_Tables(i).getTableName(erdInfo.mode) & "." & duplicateItems(k)
                Call Log.error(strTmp)
                resultCol = resultCol & strTmp & vbCrLf
            Next
        End If
    Next
    
    If Util.duplicatedCheck(tables, duplicateItems) Then
        result = False
        For k = LBound(duplicateItems) To UBound(duplicateItems)
            strTmp = "[表名の重複] " & duplicateItems(k)
            Call Log.error(strTmp)
            resultTab = resultTab & strTmp & vbCrLf
        Next
    End If
    
    msg = msg & vbCrLf & resultTab & vbCrLf & resultCol

    validateDuplicateTabAndCol = result
End Function
'
' データ型のチェック
'
Public Function validateDataType(ByRef msg, ByRef m_Tables() As ErTable, ByRef erdInfo As ERDInformation) As Boolean
    Dim result As Boolean
    
    Dim i As Integer
    Dim j As Integer
    
    Dim colCnt      As Integer
    Dim tmpDataType As String
    Dim tmpDataLen  As String
    Dim strTmp      As String
    
    result = True
    
    For i = LBound(m_Tables) To UBound(m_Tables)
        colCnt = m_Tables(i).getColCount
        For j = 1 To colCnt
            tmpDataType = m_Tables(i).getColumn(j - 1).getDataType()
            tmpDataLen = m_Tables(i).getColumn(j - 1).getDataLength()
        
            If Not isValidDataLength(tmpDataLen) Then
                strTmp = "[不正なデータ長] <" & tmpDataLen & ">" & m_Tables(i).getTableName(erdInfo.mode) _
                    & "." & m_Tables(i).getColumn(j - 1).getColName(erdInfo.mode)
                Call Log.error(strTmp)
                msg = msg & vbCrLf & strTmp
                result = False
            End If
            
        Next
    Next
    
    msg = msg
    validateDataType = result
    
End Function
'
'
'
Private Function isValidDataLength(ByVal dataLength As String) As Boolean
    
    If Util.isBlank(dataLength) Then
        isValidDataLength = True
        Exit Function
    End If

    If IsNumeric(dataLength) Then
        isValidDataLength = True
        Exit Function
    End If
    
    Dim lens() As String
    Dim i As Integer
    lens = Split(dataLength, ",")
    For i = LBound(lens) To UBound(lens)
        If Not IsNumeric(dataLength) Then
            isValidDataLength = False
            Exit Function
        End If
    Next
    
    isValidDataLength = True
    
End Function
