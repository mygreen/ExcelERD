VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileUtil"
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
' File関連の関数をラップしたUtilityクラス
'
'
'

Private Const INVALID_FILENO As Integer = -1

' 処理対象のファイル名を保持
Private m_Filename  As String

' 処理中のファイル番号を保持
Private m_FileNo    As Integer
'
'
'
Private Sub Class_Initialize()
    Call setFileNo(INVALID_FILENO)
End Sub

'
'
'
Private Sub Class_Terminate()
    Call closeFile
End Sub
'
'
'
Public Function isOpen() As Boolean
    isOpen = (getFileNo > 0)
End Function
'
'
'
Public Function openFile(fileName As String, mode As FileMode) As Boolean
    On Error GoTo errhandler
    openFile = True
    Call setFilename(fileName)
    Call setFileNo(FreeFile)
    
    Select Case mode
        Case FileMode.AppendMode
            Open getFilename For Append As getFileNo
            
        Case FileMode.BinaryMode
            Open getFilename For Binary As getFileNo
            
        Case FileMode.InputMode
            Open getFilename For Input As getFileNo
            
        Case FileMode.OutputMode
            Open getFilename For Output As getFileNo
            
        Case FileMode.Random
            Open getFilename For Random As getFileNo
            
        Case Else
            openFile = False
    End Select
    
    Exit Function
errhandler:
    openFile = False
End Function
'
'
'
Public Sub println(str As String)
    Print #getFileNo(), str
End Sub
'
'
'
Public Sub closeFile()
    Close getFileNo()
End Sub

'
'
'
Private Sub setFilename(newVal As String)
    m_Filename = newVal
End Sub
'
'
'
Public Function getFilename() As String
    getFilename = m_Filename
End Function
'
'
'
Private Sub setFileNo(newVal As Integer)
    m_FileNo = newVal
End Sub
'
'
'
Public Function getFileNo() As Integer
    getFileNo = m_FileNo
End Function
