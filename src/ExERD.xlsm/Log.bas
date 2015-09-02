Attribute VB_Name = "Log"
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
' Logファイルへの静的アクセスを提供
'
'
Private fu As FileUtil
Private m_IsOpen As Boolean

Public Sub initialLog(fileName As String)
    m_IsOpen = False
    Set fu = New FileUtil
    
    m_IsOpen = fu.openFile(fileName, FileMode.AppendMode)
End Sub
'
'
'
Public Sub terminateLog()
    Call fu.closeFile
    m_IsOpen = False
    Set fu = Nothing
End Sub

'
'
'
Public Sub info(str As String)
    If Not isWritable() Then
        Exit Sub
    End If

    Call fu.println(getTimeStamp & "[INFO] " & str)
End Sub
'
'
'
Public Sub error(str As String)
    If Not isWritable() Then
        Exit Sub
    End If

    Call fu.println(getTimeStamp & "[ERROR] " & str)
End Sub
'
'
'
Private Function isWritable() As Boolean
    isWritable = (m_IsOpen And fu.isOpen)
End Function
'
'
'
Private Function getTimeStamp() As String
    getTimeStamp = Date & " " & Time
End Function

