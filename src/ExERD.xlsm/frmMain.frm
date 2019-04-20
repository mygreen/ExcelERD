VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "AppTitle"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmMain"
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


Private prop        As Settings         '�ݒ�l�ێ�
Private isRunning   As Boolean

'
' ��ʏ�������
'
Private Sub UserForm_Initialize()
    
    Set prop = New Settings
    
    Call setControlTips
    Call prop.loadSettings(ThisWorkbook, False)
    Call setFormatSettings(prop)
    Call setDDLSettings(prop)
    Call setDDLOutputSettings(prop)
    Call setOdbcSettings(prop)
    Call setUISettings(prop)
    
    Me.Caption = Constants.getAppInfo()
    Call setAppAbout

End Sub
'
' About Box �̏���\������
'
Private Sub setAppAbout()
    
    lblAppName.Caption = Constants.getAppTitle()
    lblVersion.Caption = "Version " & Constants.getAppVersion()
    lblCopyRight.Caption = Constants.APP_COPY_RIGHT
    lblInfoTo.Caption = Constants.APP_AUTHOR_MAIL
    lblInfoTo2.Caption = Constants.APP_AUTHOR_MAIL2

End Sub
'
'
'
Private Sub lockControls(isLock As Boolean)
    isRunning = isLock
    Call Constants.setGlobalCancelFlag(False)
    multiExcelErdPage.Enabled = Not isLock
    
End Sub
'
' DDL�V�[�g�̃w�b�_��}������
'
Private Sub cmdInsertDDLHead_Click()
    
    Dim wrkBook     As Workbook
    Dim sheetInfo   As SheetInformation
    Dim ddlSheet    As Worksheet
    Dim errResult   As Long
    
    Call setStetusMsg("�e�[�u����`�w�b�_�[�}�������")
    Call Constants.clearSheetInfo(sheetInfo)
    Call frmSelSheetDlg.showDDLHeaderSheetSelectDialog(sheetInfo)
        
    errResult = validateSelectedSheet(sheetInfo)
    If errResult <> Util.NO_ERROR Then
        Call Util.showDialog(errResult)
        Exit Sub
    End If
    
    If sheetInfo.isNewSheet Then
        Set wrkBook = Application.Workbooks(sheetInfo.bookName)
        Set ddlSheet = wrkBook.Sheets.ADD
    Else
        Set ddlSheet = Application.Workbooks( _
                        sheetInfo.bookName).Worksheets(sheetInfo.sheetName)
    End If
    
    Call insertDDLHeader(ddlSheet)
    Call setStetusMsg("")
    
End Sub
'
' �����ݒ�̏�����
'
Private Sub cmdInitFormatSettings_Click()
    Dim tmpProp As New Settings
    
    Call setStetusMsg("�����l�Đݒ蒆���")
    Call tmpProp.loadFormatDefaultSettings
    Call setFormatSettings(tmpProp)
    Set tmpProp = Nothing
        
    Call setApplicateFormatStatus(False)
    Call setStetusMsg("")

End Sub

'
' DDL�ݒ�̏����l��ǂݍ���
'
Private Sub cmdInitDDLSettings_Click()
    Dim tmpProp As New Settings
    
    Call setStetusMsg("�����l�Đݒ蒆���")
    Call tmpProp.loadDDLDefaultSettings
    Call setDDLSettings(tmpProp)
    Set tmpProp = Nothing
    
    Call setApplicateDDLStatus(False)
    Call setStetusMsg("")

End Sub
'
' DDL�ݒ�̓K�p
'
Private Sub cmdAppDDLSetting_Click()
    Call setStetusMsg("�ݒ�K�p�����")
    If Not validateDDLSettings() Then
        Exit Sub
    End If
    Call applicateDDLSettings
    Call prop.saveDDLSettings(ThisWorkbook)
    Call setDDLSettings(prop)

    Call setStetusMsg("")
End Sub
'
' �����ݒ�K�p
'
Private Sub cmdAppFormatSetting_Click()
    Call setStetusMsg("�ݒ�K�p�����")
    If Not validateFormatSettings() Then
        Exit Sub
    End If
    Call applicateFormatSettings
    Call prop.saveFormatSettings(ThisWorkbook)
    Call setFormatSettings(prop)
    Call setStetusMsg("")
End Sub
'
' DDL�ݒ萮�����`�F�b�N
'
Private Function validateDDLSettings() As Boolean
    
    validateDDLSettings = False
    
    If strtoInt(txtStartRow.text) <= 0 Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Constants.START_ROW)
        Exit Function
    End If
    
    If strtoInt(txtStartRow.text) <= 1 Then
        Call Util.showDialog(Constants.ERR_REQUIRED_MORE_VAL, Array(Constants.START_ROW, "'2'"))
        Exit Function
    End If
    
    validateDDLSettings = True
    
End Function
'
' �����ݒ萮�����`�F�b�N
'
Private Function validateFormatSettings() As Boolean
    
    validateFormatSettings = False
    
    If strtoInt(txtFontSize.text) < 4 Or strtoInt(txtFontSize.text) > 41 Then
        Call Util.showDialog(Constants.ERR_REQUIRED_RANGE, Array("�t�H���g�T�C�Y", "'4'", "'40'"))
        Exit Function
    End If
    
    validateFormatSettings = True
    
End Function
'
' DDL�o�͐ݒ�`�F�b�N
'
Private Function validateDDLOutputSettings() As Boolean
    
    validateDDLOutputSettings = False
    
    If Util.isBlank(txtDDLPath.text) Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Array("�o�͐�t�H���_"))
        Exit Function
    End If
    
    If Util.isBlank(txtDDLFile.text) Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Array("�t�@�C����"))
        Exit Function
    End If
    
    'add tatsuo:�R�����g������A��؂蕶���̃`�F�b�N
    If Util.isBlank(txtSepStr.text) Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Array("�R�����g�����񕶎�"))
        Exit Function
    End If
    
    If Util.isBlank(txtComment.text) Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Array("��ؕ���"))
        Exit Function
    End If
    
    validateDDLOutputSettings = True
    
End Function
'
'
'
Private Function validateOutputPath(ByRef outputPath As String) As Boolean
    validateOutputPath = False
    
    On Error GoTo errhandler
    
    Dim dirStr As String
    Dim dirAttr, msg As String
    
    outputPath = Util.getCurrentPath(outputPath)
    
    dirStr = Dir$(outputPath, vbDirectory)
    If dirStr = "" Or dirStr = "." Then
        If Util.showDialog(Constants.Q_YN_CREATE_DIR, Array(outputPath)) = vbYes Then
            Call MkDir(outputPath)
        Else
            Exit Function
        End If
    Else
        ' check writable
        On Error GoTo checkFolderError
        
        If Not Util.isWritableFolder(outputPath) Then
            Call Util.showDialog(Constants.ERR_NOT_WRITE_FOLDER, Array(outputPath))
            validateOutputPath = False
            Exit Function
        End If
        
    End If
    
    validateOutputPath = True
    
    Exit Function
errhandler:
    Call Util.showDialog(Constants.ERR_GENERAL, Array(Err.Description))
    Exit Function
        
checkFolderError:
    Call Util.showDialog(Constants.ERR_NOT_WRITE_FOLDER, Array(outputPath))
    validateOutputPath = False
    Exit Function
    
End Function
'
'
'
Private Function validateAlreadyExistsFile(fileName As String) As Boolean
    
    Dim FSO, outFile As Object

    validateAlreadyExistsFile = False

    If Dir$(fileName) <> "" Then
        If Util.showDialog(Constants.Q_YN_OVERWRITE_FILE, Array(fileName)) = vbYes Then
            'OK Over Write
        Else
            Exit Function
        End If
    End If
    
    ' check write file
    On Error GoTo checkFileError
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CreateTextFile(fileName, True).Close

    validateAlreadyExistsFile = True
    Exit Function
    
checkFileError:
    Call Util.showDialog(Constants.ERR_NOT_WRITE_FILE, Array(fileName))
    validateAlreadyExistsFile = False
    Exit Function
    
End Function

'
' DDL�ݒ��K�p
'
Private Sub applicateDDLSettings()
    
    Call prop.setStartRow(strtoInt(txtStartRow.text))
    Call prop.setColposObjectType(strtoInt(txtObjectType.text))
    Call prop.setColposLogicalTableName(strtoInt(txtLogicalTableName.text))
    Call prop.setColposPhysicalTableName(strtoInt(txtPhysicalTableName.text))
    Call prop.setColposColId(strtoInt(txtColID.text))
    Call prop.setColposLogicaColName(strtoInt(txtLogicalColName.text))
    Call prop.setColposPhysicalColName(strtoInt(txtPhysicalColName.text))
    Call prop.setColposDataType(strtoInt(txtDataType.text))
    Call prop.setColposDataLength(strtoInt(txtDataLength.text))
    Call prop.setColposNotNull(strtoInt(txtNotNull.text))
    Call prop.setColposPrimaryKey(strtoInt(txtPrimaryKey.text))
    Call prop.setColposForeingKey(strtoInt(txtForeignKey.text))
    Call prop.setColposDependenceTableName(strtoInt(txtDependTableName.text))
    Call prop.setColposRelationType(strtoInt(txtRelationType.text))
    Call prop.setColposDefaultValue(strtoInt(txtDefaultVal.text))

End Sub
'
' �����ݒ��K�p
'
Private Sub applicateFormatSettings()
    
    Call prop.setFontSize(strtoInt(txtFontSize.text))
    Call prop.setMarginLeft(strtoInt(txtMarginLeft.text))
    Call prop.setMarginTop(strtoInt(txtMarginTop.text))
    Call prop.setInterval(strtoInt(txtInterval.text))
    Call prop.setWidthlimit(strtoInt(txtWidhthLimit.text))

End Sub
'
' ���[�U�C���^�[�t�F�[�X�ݒ�l��K�p
'
Private Sub applicateUISetting()
    
    ' �o�̓��f��
    Dim modelmode As Integer
    modelmode = 0
    If optLogicalModel.Value = True Then
        modelmode = 1
    ElseIf optPhisicaAndLogicalModel.Value = True Then
        modelmode = 2
    End If
    Call prop.setModelMode(modelmode)
    
    ' �o�̓I�v�V���� - �����[�V����
    Dim outRel As Integer
    outRel = 1
    If chkRelation.Value = False Then
        outRel = 0
    End If
    Call prop.setOutputRelation(outRel)
    
    ' �o�̓I�v�V���� - ��L�[�Ȃǂ̏ȗ�
    Dim outEli As Integer
    outEli = 1
    If chkElision.Value = False Then
        outEli = 0
    End If
    Call prop.setOutputElision(outEli)
    
    
End Sub
'
' ERD�쐬
'
Private Sub cmdCreateErd_Click()
    
    Dim wrkBook     As Workbook
    Dim sheetInfo   As SheetInformation
    Dim erdInfo     As ERDInformation
    Dim erdDoc      As New ExERDDocument
    Dim ddlSheet    As Worksheet
    Dim erdSheet    As Worksheet
    Dim mode        As ERDMODE
    Dim errResult   As Long
    
    Call saveUISettings
    Call lockControls(True)
    
    ' ---------- DDL Sheet Select ----------
    Call Constants.clearSheetInfo(sheetInfo)
    Call frmSelSheetDlg.showDDLSheetSelectDialog(sheetInfo)
        
    errResult = validateSelectedSheet(sheetInfo)
    If errResult <> Util.NO_ERROR Then
        Call Util.showDialog(errResult)
        GoTo finally
    End If
    Set ddlSheet = Application.Workbooks( _
                    sheetInfo.bookName).Worksheets(sheetInfo.sheetName)
    
    ' ---------- ERD Sheet Select ----------
    Call Constants.clearSheetInfo(sheetInfo)
    Call frmSelSheetDlg.showERDSheetSelectDialog(sheetInfo)
    errResult = validateSelectedSheet(sheetInfo)
    If errResult <> Util.NO_ERROR Then
        Call Util.showDialog(errResult)
        GoTo finally
    End If
    
    If sheetInfo.isNewSheet Then
        Set wrkBook = Application.Workbooks(sheetInfo.bookName)
        Set erdSheet = wrkBook.Sheets.ADD
    Else
        Set erdSheet = Application.Workbooks( _
                        sheetInfo.bookName).Worksheets(sheetInfo.sheetName)
    End If
    
    '---------- Create ERD ----------
    If MsgBox("'" & ddlSheet.Name & "'��ER�}�� '" & _
                    erdSheet.Name & "'�֍쐬���܂��B", _
                    vbYesNo Or vbQuestion, Constants.getAppInfo) = vbYes Then
        
        Call setStetusMsg("ERD �쐬�����")
        
        If erdDoc.loadTableData(ddlSheet) <> -1 Then
            
            If optLogicalModel.Value = True Then
                erdInfo.mode = Logical
            ElseIf optPhisicalModel = True Then
                erdInfo.mode = Physical
            ElseIf optPhisicaAndLogicalModel = True Then
                erdInfo.mode = PhysicalAndLogical
            End If
            erdInfo.fontSize = prop.getFontSize
            
            Call erdDoc.drawERD(erdSheet, erdInfo)
        End If
    End If
    
    Call MsgBox("�I�����܂����B", vbInformation, Constants.getAppInfo)
finally:
    Call setStetusMsg("")
    Call lockControls(False)
End Sub
'
' ���o�[�X���s
'
Private Sub cmdReverse_Click()
    Dim dataTypeFile As String
    
    dataTypeFile = txtDatatypeFile.text
    
    Call setStetusMsg("ERD �쐬��(���o�[�X�G���W�j�A�����O) ���")
    Call frmReverse.showReverseDialog(prop, dataTypeFile)
    txtDatatypeFile.text = dataTypeFile
    
    Call setStetusMsg("")
    
    Call saveOdbcSettings
    
End Sub
'
' DDL���o�͂���
'
Private Sub cmdWriteDDL_Click()
    
    Dim wrkBook     As Workbook
    Dim sheetInfo   As SheetInformation
    Dim erdInfo     As ERDInformation
    Dim erdDoc      As New ExERDDocument
    Dim ddlSheet    As Worksheet
    Dim mode        As ERDMODE
    Dim errResult   As Long
    Dim outPath     As String
    
    ' add tastuo
    Dim ddlInfo As DDLInformation
    Dim sepStr  As String
    Dim commentStr As String
    
    ' add tatsuo
    Call saveDDLOutputSetting
    Call lockControls(True)
    
    ' ---------- Output File Select ----------
    If Not validateDDLOutputSettings() Then
        GoTo finally
    End If
    
    outPath = RTrim$(txtDDLPath.text)
    If Not validateOutputPath(outPath) Then
        GoTo finally
    End If
    
    outPath = Util.getPath(outPath) & RTrim$(txtDDLFile.text)
    If Not validateAlreadyExistsFile(outPath) Then
        GoTo finally
    End If
    
    ' ---------- DDL Sheet Select ----------
    Call Constants.clearSheetInfo(sheetInfo)
    Call frmSelSheetDlg.showDDLSheetSelectDialog(sheetInfo)
        
    errResult = validateSelectedSheet(sheetInfo)
    If errResult <> Util.NO_ERROR Then
        Call Util.showDialog(errResult)
        GoTo finally
    End If
    Set ddlSheet = Application.Workbooks( _
                    sheetInfo.bookName).Worksheets(sheetInfo.sheetName)
    
    'add tatsuo: DDL�̋�؂蕶���Ȃǂ̐ݒ�l
    Call Constants.clearDDLInfo(ddlInfo)
    sepStr = RTrim$(txtSepStr.text)
    commentStr = RTrim$(txtComment.text)
    ddlInfo.sepStr = sepStr
    ddlInfo.commentStr = commentStr
    
    
    '---------- Create DDL ----------
    If MsgBox("'" & ddlSheet.Name & "'��DDL��`�� '" & _
                    outPath & "'�֏o�͂��܂��B", _
                    vbYesNo Or vbQuestion, Constants.getAppInfo) = vbYes Then
        
        Call setStetusMsg("DDL �o�͒����")
        If erdDoc.loadTableData(ddlSheet) <> -1 Then
            Call erdDoc.writeDDL(outPath, ddlInfo)
        End If
        
    End If
    
    ' add tatsuo: DDL�t�@�C�����m�[�g�p�b�h�ŊJ��
    If chkDDLOpenWithNotepad Then
    
        If MsgBox("�I�����܂����B" & vbCrLf & "�������ŊJ���܂����H", vbInformation Or vbYesNo, Constants.getAppInfo) _
            = vbYes Then
            Call Util.openFileWithNotepad(outPath)
        End If
    
    Else
        ' add tatsuo:�t�@�C����OS�Ɋ֘A�t����ꂽ�A�v�����[�V�����ŊJ���B
        If MsgBox("�I�����܂����B" & vbCrLf & "OS�Ɋ֘A�t�����ꂽ�A�v���P�[�V�����ŊJ���܂����H", vbInformation Or vbYesNo, Constants.getAppInfo) _
            = vbYes Then
             Call Util.openFileWithOS(outPath)
        End If
    
    End If
    
    Call setStetusMsg("")
    
finally:
    Call lockControls(False)
End Sub
'
' DDL�ݒ����ʂɔ��f
'
Private Sub setDDLSettings(newProp As Settings)
    
    Call setStetusMsg("�V�[�g�ʒu���ݒ蒆���")
    With newProp
        Call setNumTextField(txtStartRow, .getStartRow())
        Call setNumTextField(txtObjectType, .getColposObjectType())
        Call setNumTextField(txtLogicalTableName, .getColposLogicalTableName())
        Call setNumTextField(txtPhysicalTableName, .getColposPhysicalTableName())
        Call setNumTextField(txtColID, .getColposColId())
        Call setNumTextField(txtLogicalColName, .getColposLogicaColName())
        Call setNumTextField(txtPhysicalColName, .getColposPhysicalColName())
        Call setNumTextField(txtDataType, .getColposDataType())
        Call setNumTextField(txtDataLength, .getColposDataLength())
        Call setNumTextField(txtNotNull, .getColposNotNull())
        Call setNumTextField(txtPrimaryKey, .getColposPrimaryKey())
        Call setNumTextField(txtForeignKey, .getColposForeingKey())
        Call setNumTextField(txtDependTableName, .getColposDependenceTableName())
        Call setNumTextField(txtRelationType, .getColposRelationType())
        Call setNumTextField(txtDefaultVal, .getColposDefaultValue())
    End With
    Call setApplicateDDLStatus(True)
    Call setStetusMsg("")
    
End Sub
'
' ���l���ڂ��e�L�X�g�ɃZ�b�g
'
Private Sub setNumTextField(ByRef txtField As Variant, intValue As Integer, Optional baseValue)
    
    If IsMissing(baseValue) Then
        baseValue = 1
    End If
    
    If TypeName(txtField) = "TextBox" Then
        If intValue < baseValue Then
            txtField.text = ""
            txtField.BackColor = &H8000000F
        Else
            txtField.text = CStr(intValue)
            txtField.BackColor = &H80000005
        End If
    End If

End Sub
'
' �����ݒ����ʂɃZ�b�g
'
Private Sub setFormatSettings(newProp As Settings)
    
    Call setStetusMsg("�����ݒ蒆���")
    With newProp
        Call setNumTextField(txtFontSize, .getFontSize())
        Call setNumTextField(txtMarginTop, .getMarginTop(), 0)
        Call setNumTextField(txtMarginLeft, .getMarginLeft(), 0)
        Call setNumTextField(txtInterval, .getInterval(), 0)
        Call setNumTextField(txtWidhthLimit, .getWidthlimit())
    End With
    Call setApplicateFormatStatus(True)
    Call setStetusMsg("")

End Sub
'
' DDL�o�͐ݒ����ʂɃZ�b�g
'
Private Sub setDDLOutputSettings(newProp As Settings)
    
    Call setStetusMsg("DDL�o�͐ݒ襥�")
    With newProp
        txtDDLPath.text = .getDDLOutputPath
        txtDDLFile.text = .getDDLOutputFile
        txtSepStr.text = .getDDLSeqStr
        txtComment.text = .getDDLCommentStr
        
        If .getDDLOutputLogicalName = 1 Then
            chkDDLLogicalName = True
        End If
        
        
        If .getDDLOpenNodepad = 1 Then
            chkDDLOpenWithNotepad = True
        End If
        
        
    End With
    Call setStetusMsg("")

End Sub
'
' ODBC�ݒ����ʂɃZ�b�g
'
Private Sub setOdbcSettings(newProp As Settings)
    
    Call setStetusMsg("ODBC�ݒ襥�")
    With newProp
        txtDatatypeFile.text = .getOdbcDatatypeFile
    End With
    Call setStetusMsg("")
    
End Sub
'
' UI�ݒ����ʂɃZ�b�g
'
Private Sub setUISettings(newProp As Settings)
    
    ' ���f���̎�ނ̐ݒ�
    If newProp.getModelMode = 1 Then
        optLogicalModel.Value = True
    ElseIf newProp.getModelMode = 2 Then
        optPhisicaAndLogicalModel = True
    Else
        optPhisicalModel.Value = True
    End If
    
    ' �o�̓I�v�V���� - �����[�V�����̏o�͂̐ݒ�
    If newProp.getOutputRelation = 0 Then
        chkRelation.Value = False
        ' DDL�̊O���L�[����Ƃ̘A��
        chkDDLRelation.Value = False
    Else
        chkRelation.Value = True
        ' DDL�̊O���L�[����Ƃ̘A��
        chkDDLRelation.Value = True
    End If
    
    ' �o�̓I�v�V���� - ��L�[�Ȃǂ̏ȗ��̐ݒ�
    If newProp.getOutputElision = 0 Then
        chkElision.Value = False
    Else
        chkElision.Value = True
    End If

    
End Sub
'
' DDL�K�p�{�^���g�p��/�s�؂�ւ�
'
Private Sub setApplicateDDLStatus(status As Boolean)
    
    cmdAppDDLSetting.Enabled = Not status

End Sub
'
' �����K�p�{�^���g�p��/�s�؂�ւ�
'
Private Sub setApplicateFormatStatus(status As Boolean)
    cmdAppFormatSetting.Enabled = Not status
End Sub

'
' �f�[�^��`�w�b�_�[���V�[�g�ɑ}��
'
Public Sub insertDDLHeader(sheet As Worksheet)
    
    Dim rowPos As Integer
    Dim colPos As Integer
    
    With prop
        rowPos = .getStartRow - 1
        
        Call setSheetCells(sheet, rowPos, _
            .getColposObjectType(), Constants.COL_OBJECT_TYPE, Constants.COMMENT_OBJECT_TYPE)
        Call setSheetCells(sheet, rowPos, _
            .getColposLogicalTableName(), Constants.COL_LOGICAL_TABLENAME, Constants.COMMENT_LOGICAL_TABLENAME)
        Call setSheetCells(sheet, rowPos, _
            .getColposPhysicalTableName(), Constants.COL_PHYSICAL_TABLENAME, Constants.COMMENT_PHYSICAL_TABLENAME)
        Call setSheetCells(sheet, rowPos, _
            .getColposColId(), Constants.COL_COLID, Constants.COMMENT_COLID)
        Call setSheetCells(sheet, rowPos, _
            .getColposLogicaColName(), Constants.COL_LOGICAL_COLNAME, Constants.COMMENT_LOGICAL_COLNAME)
        Call setSheetCells(sheet, rowPos, _
            .getColposPhysicalColName(), Constants.COL_PHYSICAL_COLNAME, Constants.COMMENT_PHYSICAL_COLNAME)
        Call setSheetCells(sheet, rowPos, _
            .getColposDataType(), Constants.COL_DATATYPE, Constants.COMMENT_DATATYPE)
        Call setSheetCells(sheet, rowPos, _
            .getColposDataLength(), Constants.COL_DATALENGTH, Constants.COMMENT_DATALENGTH)
        Call setSheetCells(sheet, rowPos, _
            .getColposNotNull(), Constants.COL_NOTNULL, Constants.COMMENT_NOTNULL)
        Call setSheetCells(sheet, rowPos, _
            .getColposPrimaryKey(), Constants.COL_PRIMARYKEY, Constants.COMMENT_PRIMARYKEY)
        Call setSheetCells(sheet, rowPos, _
            .getColposForeingKey(), Constants.COL_FOREIGNKEY, Constants.COMMENT_FOREIGNKEY)
        Call setSheetCells(sheet, rowPos, _
            .getColposDependenceTableName(), Constants.COL_DEPENDENCE_TABLENAME, Constants.COMMENT_DEPENDENCE_TABLENAME)
        Call setSheetCells(sheet, rowPos, _
            .getColposRelationType(), Constants.COL_RELATION_TYPE, Constants.COMMENT_RELATION_TYPE)
        Call setSheetCells(sheet, rowPos, _
            .getColposDefaultValue(), Constants.COL_DEFAULT_VALUE, Constants.COMMENT_DEFAULT_VALUE)
        
    End With
End Sub
'
' �V�[�g�̃Z���ɒl��ݒ�
'
Private Sub setSheetCells(sheet As Worksheet, row As Integer, col As Integer, str As String, comment As String)
    Dim rng As Range
    On Error GoTo ErrHndler
    If (row > 0) And (col > 0) Then
        sheet.Cells(row, col).Value = str
        
        Set rng = sheet.Range(sheet.Cells(row, col), sheet.Cells(row, col))
        rng.AddComment
        Call rng.comment.text(comment)
        
    End If
    Exit Sub
ErrHndler:
End Sub



'
' �u��ݾفv�{�^��
'
Private Sub cmdCancel_Click()
    If isRunning Then
        If Util.showDialog(Q_YN_CANCEL_PROC) = vbYes Then
            Call Constants.setGlobalCancelFlag(True)
        End If
    Else
        Unload Me
    End If
End Sub

'
' �R���g���[���`�b�v��ݒ�
'
Private Sub setControlTips()

    lblErdExplain.Caption = Constants.MSG_CREATE_ERD_EXPLAIN
    lblRevExplain.Caption = Constants.MSG_REVERSE_ERD_EXPLAIN
    lblSheetPosExplain.Caption = Constants.MSG_SHEET_POS_EXPLAIN
    lblFormatExplain.Caption = Constants.MSG_FORMAT_EXPLAIN
    lblDDLExplain.Caption = Constants.MSG_CREATE_DDL_EXPLAIN
    
    fraModelKind.ControlTipText = Constants.TIPS_MODEL_KIND
    chkRelation.ControlTipText = Constants.TIPS_RELATION
    chkElision.ControlTipText = Constants.TIPS_ELISION
    txtDatatypeFile.ControlTipText = Constants.TIPS_DATATYPEFILE
        
    txtStartRow.ControlTipText = Constants.TIPS_START_ROW
    txtObjectType.ControlTipText = Constants.TIPS_OBJECT_TYPE
    txtLogicalTableName.ControlTipText = Constants.TIPS_LOGICAL_TABLENAME
    txtPhysicalTableName.ControlTipText = Constants.TIPS_PHISYCAL_TABLENAME
    txtColID.ControlTipText = Constants.TIPS_COLID
    txtLogicalColName.ControlTipText = Constants.TIPS_LOGICAL_COLNAME
    txtPhysicalColName.ControlTipText = Constants.TIPS_PHISYCAL_COLNAME
    txtDataType.ControlTipText = Constants.TIPS_DATATYPE
    txtDataLength.ControlTipText = Constants.TIPS_DATALENGTH
    txtNotNull.ControlTipText = Constants.TIPS_NOTNULL
    txtPrimaryKey.ControlTipText = Constants.TIPS_PRIMARYKEY
    txtForeignKey.ControlTipText = Constants.TIPS_FOREIGNKEY
    txtDependTableName.ControlTipText = Constants.TIPS_DEPENDENCE_TABLENAME
    txtRelationType.ControlTipText = Constants.TIPS_RELATION_TYPE
    
    txtFontSize.ControlTipText = Constants.TIPS_FONTSIZE
    txtMarginLeft.ControlTipText = Constants.TIPS_MARGIN_LEFT
    txtMarginTop.ControlTipText = Constants.TIPS_MARGIN_TOP
    txtWidhthLimit.ControlTipText = Constants.TIPS_WIDTH_LIMIT
    txtInterval.ControlTipText = Constants.TIPS_INTERVAL
    
    txtDDLPath.ControlTipText = Constants.TIPS_DDL_OUTPUT_DIR
    txtDDLFile.ControlTipText = Constants.TIPS_DDL_OUTPUT_FILE
    txtSepStr.ControlTipText = Constants.TIPS_DDL_SEP_TEXT
    'add tatsuo: �R�����g�̃q���g�\��
    txtComment.ControlTipText = Constants.TIPS_DDL_COMMENT

    
End Sub
'
' ���[�U�ݒ��ۑ�
'
Private Sub saveOdbcSettings()
    
    Call prop.setOdbcDatatypeFile(txtDatatypeFile.text)
    Call prop.saveOdbcSettings(ThisWorkbook)
        
End Sub
'
' ���[�U�ݒ��ۑ�
'
Private Sub saveUISettings()
    
    Call applicateUISetting
    Call prop.saveUISettings(ThisWorkbook)
        
End Sub
'
'
'
Private Sub saveDDLOutputSetting()
    
    Dim outRel As Integer
    Dim ddlOpenNotepad As Integer
    Dim ddlOutputLogicalName As Integer
    
    outRel = 1
    If chkDDLRelation.Value = False Then
        outRel = 0
    End If
    Call prop.setOutputRelation(outRel)

    Call prop.setDDLOutputPath(RTrim$(txtDDLPath.text))
    Call prop.setDDLOutputFile(RTrim$(txtDDLFile.text))
    
    'add tatsuo:��؂蕶���A�R�����g�̕����̕ۑ�
    Call prop.setDDLSepStr(RTrim$(txtSepStr.text))
    Call prop.setDDLCommentStr(RTrim$(txtComment.text))
    
    ddlOpenNotepad = 1
    If chkDDLOpenWithNotepad.Value = False Then
        ddlOpenNotepad = 0
    End If
    Call prop.setDDLOpenNodepad(ddlOpenNotepad)
    
    ddlOutputLogicalName = 1
    If chkDDLLogicalName.Value = False Then
        ddlOutputLogicalName = 0
    End If
    Call prop.setDDLOutputLogicalName(ddlOutputLogicalName)
    
    
    Call prop.saveDDLOutputSettings(ThisWorkbook)

End Sub
'
' �X�e�[�^�X���b�Z�[�W��\��
'
Private Sub setStetusMsg(msg As String)

    lblStatus.Caption = msg
    DoEvents

End Sub
'
'
'
Private Sub cmdBrowzDir_Click()
   'txtDDLPath.text = Util.browseForFolder(Constants.MSG_DDL_OUTPUTDIR_SELECT, 0)
   txtDDLPath.text = Util.browseForFolder2(Constants.MSG_DDL_OUTPUTDIR_SELECT, 0, txtDDLPath.text)
   
End Sub
'
' DDL���J���{�^�����N���b�N�����ꍇ
'
Private Sub openDDL_Click()
    
    Dim filePath As String
    
    If Util.isBlank(txtDDLPath.text) Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Array("�o�̓t�H���_"))
        Exit Sub
    End If
    
    If Util.isBlank(txtDDLFile.text) Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Array("�t�@�C����"))
        Exit Sub
    End If
    
    filePath = Util.getCurrentPath(RTrim$(txtDDLPath.text))
    filePath = Util.getPath(filePath) & RTrim$(txtDDLFile.text)
    If Dir$(filePath) = "" Then
        Call Util.showDialog(Constants.ERR_NOT_EXIST_FILE, Array(filePath))
        Exit Sub
    End If
    
    If chkDDLOpenWithNotepad Then
        Call Util.openFileWithNotepad(filePath)
    Else
        Call Util.openFileWithOS(filePath)
    End If
    
End Sub

'
'
'
Private Sub txtStartRow_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtObjectType_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtLogicalTableName_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtPhysicalTableName_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtColID_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtLogicalColName_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtPhysicalColName_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtDataType_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtDataLength_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtNotNull_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtPrimaryKey_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtForeignKey_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtDependTableName_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtRelationType_Change()
    Call setApplicateDDLStatus(False)
End Sub
'
'
'
Private Sub txtFontSize_Change()
    Call setApplicateFormatStatus(False)
End Sub
'
'
'
Private Sub txtMarginTop_Change()
    Call setApplicateFormatStatus(False)
End Sub
'
'
'
Private Sub txtMarginLeft_Change()
    Call setApplicateFormatStatus(False)
End Sub
'
'
'
Private Sub txtInterval_Change()
    Call setApplicateFormatStatus(False)
End Sub
'
'
'
Private Sub txtWidhthLimit_Change()
    Call setApplicateFormatStatus(False)
End Sub
'
'
'
Private Sub txtDefaultVal_Change()
    Call setApplicateFormatStatus(False)
End Sub
'
'
'
Private Sub chkDDLRelation_Change()
    chkRelation.Value = chkDDLRelation.Value
End Sub
'
'
'
Private Sub chkRelation_Change()
    chkDDLRelation.Value = chkRelation.Value
End Sub
'
'
'
Private Sub cmdDatatypeFile_Click()
    Dim openFileName As Variant
    
    openFileName = Util.chooseFile(REVERSE_DATATYPEFILE_FILTER, ThisWorkbook.Path)
    If openFileName <> False Then
        txtDatatypeFile.text = openFileName
    End If
    
End Sub
'
'
'
Private Sub cmdLog_Click()
    Call Util.openFileWithNotepad(Util.getLogFilename(ThisWorkbook))
End Sub
