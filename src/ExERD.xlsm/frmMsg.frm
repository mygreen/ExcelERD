VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMsg 
   Caption         =   "MessageForm"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   OleObjectBlob   =   "frmMsg.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMsg"
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


'選択されたボタン
Private m_selectedButton As Long

'
' 画面表示モードによりコントロールを初期化する
'
Private Sub initControl(title As String, msg As String, text As String, ByVal btnId As VbMsgBoxStyle)
    Me.Caption = title
    lblMsg.Caption = msg
    txtContent.text = text
     
    If (btnId And vbCancel) <> 0 Then
        cmdCancel.Visible = True
    Else
        cmdCancel.Visible = False
    End If
    
    If (btnId And vbOK) <> 0 Then
        cmdOk.Visible = True
    Else
        cmdOk.Visible = False
    End If
    
    If (btnId And vbNo) <> 0 Then
        cmdNo.Visible = True
    Else
        cmdNo.Visible = False
    End If
    
    If (btnId And vbYes) <> 0 Then
        cmdYes.Visible = True
    Else
        cmdYes.Visible = False
    End If
    
End Sub

Public Sub showMessageWindow(title As String, msg As String, text As String, ByVal btnId As VbMsgBoxStyle, ByRef selectedButon As Long)
    Call initControl(title, msg, text, btnId)
    
    Call Me.show
    selectedButon = m_selectedButton
    Unload Me

End Sub
Private Sub cmdCancel_Click()
    m_selectedButton = vbCancel
    Me.Hide
End Sub
Private Sub cmdNo_Click()
    m_selectedButton = vbNo
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    m_selectedButton = vbOK
    Me.Hide
End Sub

Private Sub cmdYes_Click()
    m_selectedButton = vbYes
    Me.Hide
End Sub

