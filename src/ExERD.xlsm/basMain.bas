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


Private Const ERD_BUILDER       As String = "ExcelERD"      '�c�[���o�[��
Private Const BTN_ERD_BUILD     As String = "ERD�쐬"       '�{�^����
'
'ToolBar���Z�b�g
'
Private Sub loadToolBar()
   Dim cbrGatherImgs   As CommandBar
   Dim btnGetImages    As CommandBarButton
   On Error Resume Next
   ' �R�}���h �o�[�����ɑ��݂��邩�ǂ������m�F���܂��B
   ' Set cbrGatherImgs = CommandBars(ERD_BUILDER)
   ' �R�}���h �o�[�����݂��Ȃ��ꍇ�͍쐬���܂��B
   If cbrGatherImgs Is Nothing Then
      Err.clear
      Set cbrGatherImgs = CommandBars.ADD(ERD_BUILDER)
      ' �R�}���h �o�[��\�����܂��B
      cbrGatherImgs.Visible = True
      ' �{�^�� �R���g���[����ǉ����܂��B
      Set btnGetImages = cbrGatherImgs.Controls.ADD
      
      With btnGetImages
         .Style = msoButtonIconAndCaption
         .Caption = BTN_ERD_BUILD
         .Tag = BTN_ERD_BUILD
         ' �{�^�����N���b�N���ꂽ�Ƃ��Ɏ��s����v���V�[�W�����w�肵�܂��B
         .OnAction = "excelErMain"
         .FaceId = 270&
      End With
   Else
      ' �����̃R�}���h �o�[��\�����܂��B
      cbrGatherImgs.Visible = True
   End If
End Sub
'
'ToolBar���폜
'
Private Sub unloadToolBar()
  'On Error Resume Next
  On Error GoTo errhandler
  
   ' ���݂���R�}���h �o�[���폜���܂��B
   CommandBars(ERD_BUILDER).Delete
    
    Exit Sub
errhandler:
    'NOP
End Sub
'
'�t�@�C�����J�����Ƃ��Ɏ��s
'
Public Sub Auto_Open()
    Call loadToolBar
End Sub
'
'�t�@�C��������Ƃ��Ɏ��s
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


