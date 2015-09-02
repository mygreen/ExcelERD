VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErDDLWriter"
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
' DDL���o�͂���N���X
'
'
'

' �o��DDL �L�[���[�h�萔
Private Const CREATE_TABLE  As String = "CREATE TABLE"
Private Const NOT_NULL      As String = "NOT NULL"
Private Const REFERENCES    As String = "REFERENCES"
Private Const FOREIGN_KEY   As String = "FOREIGN KEY"
Private Const DEFAULT       As String = "DEFAULT"
Private Const ALTER_TABLE   As String = "ALTER TABLE"
Private Const ADD           As String = "ADD"

'
' �w�肳�ꂽ�t�@�C����DDL���o�͂���
'
Public Function wirteDDL(fileName As String, tables() As ErTable, ddlInfo As DDLInformation, _
                        Optional isOutputRelation)
                        
    On Error GoTo errhandler
    
    Dim fu              As New FileUtil
    
    Dim refListAry()    As ErRefTabList
    Dim refListCnt      As Integer
    Dim refExist        As Boolean
    Dim refList         As ErRefTabList
    
    Dim mode            As ERDMODE
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim selDir          As String
    
    refListCnt = 0
    ReDim Preserve refListAry(refListCnt)
    refExist = False
    
    '�t�@�C�����J��
    If Not fu.openFile(fileName, FileMode.OutputMode) Then
        Call Util.showDialog(Constants.ERR_GENERAL, "�t�@�C�����J���܂���")
        Exit Function
    End If
    
    'ERD���[�h = ����
    mode = ERDMODE.Physical
    For i = LBound(tables) To UBound(tables)
        
        Dim hasRef      As Boolean
        
        If Not tables(i).isView Then
            Dim colLen      As Integer
            Dim dLen        As Integer
            Dim tableName   As String
        
            tableName = tables(i).getTableName(mode)
            
            Set refList = New ErRefTabList
            Call refList.clear
            Call refList.setTableName(tableName)
            
            colLen = tables(i).getMaxColNameLength(mode)
            dLen = tables(i).getMaxDataTypeLength()
            
            
            If Not Util.isBlank(ddlInfo.commentStr) Then
                Call fu.println(ddlInfo.commentStr & " " & tables(i).getTableName(ERDMODE.Logical))
            End If
            
            Call Log.info(tables(i).getTableName(mode) & " ��DDL���o�͂��܂�")
            Call fu.println(CREATE_TABLE & Space(1) & tableName & Space(1) & "(")
                
            Dim pk()    As String
            Dim pkCnt   As Integer
            
            pkCnt = 0
            ReDim pk(pkCnt)  'ver 1.0.1
            
            hasRef = False
            For j = 0 To tables(i).getColCount - 1
                Dim line As String
                
                Dim col As ErColumun
                Set col = tables(i).getColumn(j)
                
                line = Util.indent(1) & _
                        IIf(j = 0, Space(1), ",") & _
                        Util.padding(col.getPhysicName(), colLen + 1) & _
                        Util.padding(col.getDataTypeString(), dLen + 1)
                
                If Not Util.isBlank(col.getDefaultValue) Then
                    line = line & Space(1) & _
                        DEFAULT & Space(1) & col.getDefaultValue
                End If
                If col.isNotNull Then
                    line = line & Space(1) & NOT_NULL
                End If
                
                Call fu.println(RTrim$(line))
                
                If col.isPk Then
                    ReDim Preserve pk(pkCnt)
                    pk(pkCnt) = col.getPhysicName()
                    pkCnt = pkCnt + 1
                End If
                If isOutputRelation And col.isFk Then
                    Call refList.ADD(col.getRefTable, col.getPhysicName, col.getRefCol)
                    hasRef = True
                End If
            Next
            If hasRef Then
                refExist = True
                ReDim Preserve refListAry(refListCnt)
                Set refListAry(refListCnt) = refList
                refListCnt = refListCnt + 1
            End If
            
            Dim pkStr As String
            Dim isFirst As Boolean
            pkStr = ""
            isFirst = True
            For k = LBound(pk) To UBound(pk)
                If Not Util.isBlank(pk(k)) Then
                    pkStr = pkStr & IIf(isFirst, "", ",") & pk(k)
                    isFirst = False
                End If
            Next
            If Not Util.isBlank(pkStr) Then
                Call fu.println(Util.indent(1) & ",PRIMARY KEY(" & pkStr & ")")
            End If
            
            Call fu.println(")")
            
            If Not Util.isBlank(ddlInfo.sepStr) Then
                Call fu.println(ddlInfo.sepStr)
            End If
            Call fu.println("")
            
        End If
    Next
    
    If isOutputRelation And refExist Then
        Dim strRef As String
        strRef = ""
        strRef = getReferenceCondition(refListAry, ddlInfo)
        If Not Util.isBlank(strRef) Then
            Call fu.println(strRef)
        End If
    End If
errhandler:
    Call fu.closeFile
    
    Dim msg As String
    If Err <> 0 Then
        Call showErrMsg("DDL�������ݎ��ɃG���[���������܂���")
    End If
    
End Function
'
'
'
Public Function getReferenceCondition(refTabListAry() As ErRefTabList, _
    ddlInfo As DDLInformation) As String
    
    Dim result     As String
    Dim refLines() As String
    Dim refCnt     As Integer
    
    Dim aryC    As Integer
    Dim i       As Integer
    Dim j       As Integer
    Dim k       As Integer
    Dim table   As ErTable
    Dim col     As ErColumun
    Dim sepChar As String
    Dim refTabList As ErRefTabList
    refCnt = 0
    ReDim refLines(refCnt)
    
    Dim strCols As String
    Dim strRefCols As String
    result = ""
    
    Dim sep As String
    sep = ""
    If Not Util.isBlank(ddlInfo.sepStr) Then
        sep = ddlInfo.sepStr
    End If
    
    For aryC = LBound(refTabListAry) To UBound(refTabListAry)
        Set refTabList = refTabListAry(aryC)
        
        For i = 0 To refTabList.getTableCount - 1
            Set table = refTabList.getTable(i)
            
            strCols = ""
            strRefCols = ""
            For j = 0 To table.getColCount - 1
                sepChar = IIf(j = 0, "", ",")
                Set col = table.getColumn(j)
                strCols = strCols & sepChar & col.getPhysicName
                strRefCols = strRefCols & sepChar & col.getRefCol
            Next
            If Not Util.isBlank(strCols) Then
            
                ReDim Preserve refLines(refCnt)
                refLines(refCnt) = ALTER_TABLE & " " & refTabList.getTableName & Space(1) & _
                    ADD & Space(1) & "(" & FOREIGN_KEY & "(" & strCols & ")" & Space(1) & _
                    REFERENCES & Space(1) & table.getPhisicName & "(" & strRefCols & "))" & vbCrLf & _
                    sep
                refCnt = refCnt + 1
            End If
        Next
        
    Next
    
    Dim isFirstLine As Boolean
    For k = LBound(refLines) To UBound(refLines)
        If Not Util.isBlank(refLines(k)) Then
            result = result & vbCrLf & refLines(k)
        End If
    Next
    
    getReferenceCondition = result
End Function