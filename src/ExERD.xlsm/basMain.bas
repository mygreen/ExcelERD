Attribute VB_Name = "basMain"
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


Private Const ERD_BUILDER       As String = "ExcelERD"      'ツールバー名
Private Const BTN_ERD_BUILD     As String = "ERD作成"       'ボタン名
'
'ToolBarをセット
'
Private Sub loadToolBar()
   Dim cbrGatherImgs   As CommandBar
   Dim btnGetImages    As CommandBarButton
   On Error Resume Next
   ' コマンド バーが既に存在するかどうかを確認します。
   ' Set cbrGatherImgs = CommandBars(ERD_BUILDER)
   ' コマンド バーが存在しない場合は作成します。
   If cbrGatherImgs Is Nothing Then
      Err.clear
      Set cbrGatherImgs = CommandBars.ADD(ERD_BUILDER)
      ' コマンド バーを表示します。
      cbrGatherImgs.Visible = True
      ' ボタン コントロールを追加します。
      Set btnGetImages = cbrGatherImgs.Controls.ADD
      
      With btnGetImages
         .Style = msoButtonIconAndCaption
         .Caption = BTN_ERD_BUILD
         .Tag = BTN_ERD_BUILD
         ' ボタンがクリックされたときに実行するプロシージャを指定します。
         .OnAction = "excelErMain"
         .FaceId = 270&
      End With
   Else
      ' 既存のコマンド バーを表示します。
      cbrGatherImgs.Visible = True
   End If
End Sub
'
'ToolBarを削除
'
Private Sub unloadToolBar()
  'On Error Resume Next
  On Error GoTo errhandler
  
   ' 存在するコマンド バーを削除します。
   CommandBars(ERD_BUILDER).Delete
    
    Exit Sub
errhandler:
    'NOP
End Sub
'
'ファイルを開いたときに実行
'
Public Sub Auto_Open()
    Call loadToolBar
End Sub
'
'ファイルを閉じたときに実行
'
Public Sub Auto_Close()
    Call unloadToolBar
End Sub
Public Sub excelErMain()
    Dim fm As New frmMain
    On Error GoTo MailErrHandler
    
    fm.Show
    
    Exit Sub
MailErrHandler:
    MsgBox Err.Description, vbCritical, Constants.getAppInfo()
End Sub


