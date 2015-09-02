VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelSheetDlg 
   Caption         =   "Sheet Selection Dialog"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   OleObjectBlob   =   "frmSelSheetDlg.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSelSheetDlg"
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


' ��ʕ\�����[�h
Private Enum ShowMode
    DDL_SELECTION = 0&          'DDL�V�[�g
    ERD_SELECTEION = 1&         'ERD�V�[�g
    DDL_HEAD_SELECTION = 2&     'DDL�w�b�_�}���V�[�g
End Enum

' �I���V�[�g���
Private m_sheetInfo As SheetInformation

' ��ʕ\�����[�h
Private m_mode As ShowMode

'
' ��ʂ̏�����
'
Private Sub initForm()
    Dim i As Integer
    
    For i = 1 To Application.Workbooks.Count
        Call cmbBooks.AddItem(Application.Workbooks(i).Name, i - 1)
    Next

End Sub
'
' ��ʕ\�����[�h�ɂ��R���g���[��������������
'
Private Sub initControl(mode As ShowMode)
    m_mode = mode
    
    Select Case mode
    Case ShowMode.DDL_SELECTION
        Me.Caption = Constants.TITLE_DDL_SEL_SHEET
        lblSheetSelMsg.Caption = Constants.MSG_DDL_SHEET_SELECT
    Case ShowMode.ERD_SELECTEION
        Me.Caption = Constants.TITLE_ERD_SEL_SHEET
        lblSheetSelMsg.Caption = Constants.MSG_ERD_SHEET_SELECT
    Case ShowMode.DDL_HEAD_SELECTION
        Me.Caption = Constants.TITLE_DDL_HEAD_SHEET
        lblSheetSelMsg.Caption = Constants.MSG_DDL_HEAD_SHEET_SELECT
    End Select

    Dim i As Integer
    For i = 1 To Application.Workbooks.Count
        Call cmbBooks.AddItem(Application.Workbooks(i).Name, i - 1)
    Next
    
    If Application.Workbooks.Count > 0 Then
        If Application.Workbooks.Count >= 2 Then
            cmbBooks.text = Application.Workbooks(2).Name
        Else
            cmbBooks.text = Application.Workbooks(1).Name
        End If
    End If
    
End Sub
'
' DDL�V�[�g�I����ʂ�\������
'  �I�����ꂽ���ʂ� sheetInfo �ɐݒ肷��
'
Public Sub showDDLSheetSelectDialog(ByRef sheetInfo As SheetInformation)
    
    Call initControl(ShowMode.DDL_SELECTION)
    m_sheetInfo.mode = SheetMode.DDL
    Call Me.Show(vbModal)
    sheetInfo = m_sheetInfo
    Unload Me

End Sub
'
' DDL�w�b�_�}���V�[�g�I����ʂ�\������
'  �I�����ꂽ���ʂ� sheetInfo �ɐݒ肷��
'
Public Sub showDDLHeaderSheetSelectDialog(ByRef sheetInfo As SheetInformation)
    
    Call initControl(ShowMode.DDL_HEAD_SELECTION)
    m_sheetInfo.mode = SheetMode.DDL_HEAD
    Call Me.Show(vbModal)
    sheetInfo = m_sheetInfo
    Unload Me

End Sub
'
' ERD�o�͐�V�[�g�I����ʂ�\������
'  �I�����ꂽ���ʂ� sheetInfo �ɐݒ肷��
'
Public Sub showERDSheetSelectDialog(ByRef sheetInfo As SheetInformation)
    
    Call initControl(ShowMode.ERD_SELECTEION)
    m_sheetInfo.mode = SheetMode.ERD
    Call Me.Show(vbModal)
    sheetInfo = m_sheetInfo
    Unload Me

End Sub
'
' Workbook��I��
'
Private Sub cmbBooks_Change()
    Dim book As Workbook
    Dim i As Integer
    
    m_sheetInfo.bookName = cmbBooks.text
    
    Set book = Application.Workbooks(m_sheetInfo.bookName)
    
    lstSheets.clear
    For i = 1 To book.Worksheets.Count
        Call lstSheets.AddItem(book.Worksheets(i).Name, i - 1)
    Next
    
    If (m_mode = ShowMode.ERD_SELECTEION) Or (m_mode = ShowMode.DDL_HEAD_SELECTION) Then
        Call lstSheets.AddItem(Constants.FIELD_NEW_SHEET, i - 1)
    End If
    
End Sub
'
' �L�����Z��
'
Private Sub cmdCancel_Click()
    
    m_sheetInfo.selected = CommandCondition.CANCELL
    Me.Hide
    
End Sub
'
' OK
'
Private Sub cmdOk_Click()
    
    m_sheetInfo.selected = CommandCondition.OK
    Me.Hide
    
End Sub
'
' Worksheet ��ύX
'
Private Sub lstSheets_Change()
    
    m_sheetInfo.sheetName = lstSheets.text
    If (m_sheetInfo.mode = SheetMode.ERD) Or (m_sheetInfo.mode = SheetMode.DDL_HEAD) Then
        If lstSheets.ListCount = lstSheets.ListIndex + 1 Then
            m_sheetInfo.isNewSheet = True
        End If
    End If
    
End Sub

