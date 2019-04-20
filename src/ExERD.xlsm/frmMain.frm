VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "AppTitle"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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


Private prop        As Settings         '設定値保持
Private isRunning   As Boolean

'
' 画面初期化時
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
' About Box の情報を表示する
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
' DDLシートのヘッダを挿入する
'
Private Sub cmdInsertDDLHead_Click()
    
    Dim wrkBook     As Workbook
    Dim sheetInfo   As SheetInformation
    Dim ddlSheet    As Worksheet
    Dim errResult   As Long
    
    Call setStetusMsg("テーブル定義ヘッダー挿入中･･･")
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
' 書式設定の初期化
'
Private Sub cmdInitFormatSettings_Click()
    Dim tmpProp As New Settings
    
    Call setStetusMsg("初期値再設定中･･･")
    Call tmpProp.loadFormatDefaultSettings
    Call setFormatSettings(tmpProp)
    Set tmpProp = Nothing
        
    Call setApplicateFormatStatus(False)
    Call setStetusMsg("")

End Sub

'
' DDL設定の初期値を読み込む
'
Private Sub cmdInitDDLSettings_Click()
    Dim tmpProp As New Settings
    
    Call setStetusMsg("初期値再設定中･･･")
    Call tmpProp.loadDDLDefaultSettings
    Call setDDLSettings(tmpProp)
    Set tmpProp = Nothing
    
    Call setApplicateDDLStatus(False)
    Call setStetusMsg("")

End Sub
'
' DDL設定の適用
'
Private Sub cmdAppDDLSetting_Click()
    Call setStetusMsg("設定適用中･･･")
    If Not validateDDLSettings() Then
        Exit Sub
    End If
    Call applicateDDLSettings
    Call prop.saveDDLSettings(ThisWorkbook)
    Call setDDLSettings(prop)

    Call setStetusMsg("")
End Sub
'
' 書式設定適用
'
Private Sub cmdAppFormatSetting_Click()
    Call setStetusMsg("設定適用中･･･")
    If Not validateFormatSettings() Then
        Exit Sub
    End If
    Call applicateFormatSettings
    Call prop.saveFormatSettings(ThisWorkbook)
    Call setFormatSettings(prop)
    Call setStetusMsg("")
End Sub
'
' DDL設定整合性チェック
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
' 書式設定整合性チェック
'
Private Function validateFormatSettings() As Boolean
    
    validateFormatSettings = False
    
    If strtoInt(txtFontSize.text) < 4 Or strtoInt(txtFontSize.text) > 41 Then
        Call Util.showDialog(Constants.ERR_REQUIRED_RANGE, Array("フォントサイズ", "'4'", "'40'"))
        Exit Function
    End If
    
    validateFormatSettings = True
    
End Function
'
' DDL出力設定チェック
'
Private Function validateDDLOutputSettings() As Boolean
    
    validateDDLOutputSettings = False
    
    If Util.isBlank(txtDDLPath.text) Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Array("出力先フォルダ"))
        Exit Function
    End If
    
    If Util.isBlank(txtDDLFile.text) Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Array("ファイル名"))
        Exit Function
    End If
    
    'add tatsuo:コメント文字列、区切り文字のチェック
    If Util.isBlank(txtSepStr.text) Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Array("コメント文字列文字"))
        Exit Function
    End If
    
    If Util.isBlank(txtComment.text) Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Array("区切文字"))
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
' DDL設定を適用
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
' 書式設定を適用
'
Private Sub applicateFormatSettings()
    
    Call prop.setFontSize(strtoInt(txtFontSize.text))
    Call prop.setMarginLeft(strtoInt(txtMarginLeft.text))
    Call prop.setMarginTop(strtoInt(txtMarginTop.text))
    Call prop.setInterval(strtoInt(txtInterval.text))
    Call prop.setWidthlimit(strtoInt(txtWidhthLimit.text))

End Sub
'
' ユーザインターフェース設定値を適用
'
Private Sub applicateUISetting()
    
    ' 出力モデル
    Dim modelmode As Integer
    modelmode = 0
    If optLogicalModel.Value = True Then
        modelmode = 1
    ElseIf optPhisicaAndLogicalModel.Value = True Then
        modelmode = 2
    End If
    Call prop.setModelMode(modelmode)
    
    ' 出力オプション - リレーション
    Dim outRel As Integer
    outRel = 1
    If chkRelation.Value = False Then
        outRel = 0
    End If
    Call prop.setOutputRelation(outRel)
    
    ' 出力オプション - 主キーなどの省略
    Dim outEli As Integer
    outEli = 1
    If chkElision.Value = False Then
        outEli = 0
    End If
    Call prop.setOutputElision(outEli)
    
    
End Sub
'
' ERD作成
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
    If MsgBox("'" & ddlSheet.Name & "'のER図を '" & _
                    erdSheet.Name & "'へ作成します。", _
                    vbYesNo Or vbQuestion, Constants.getAppInfo) = vbYes Then
        
        Call setStetusMsg("ERD 作成中･･･")
        
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
    
    Call MsgBox("終了しました。", vbInformation, Constants.getAppInfo)
finally:
    Call setStetusMsg("")
    Call lockControls(False)
End Sub
'
' リバース実行
'
Private Sub cmdReverse_Click()
    Dim dataTypeFile As String
    
    dataTypeFile = txtDatatypeFile.text
    
    Call setStetusMsg("ERD 作成中(リバースエンジニアリング) ･･･")
    Call frmReverse.showReverseDialog(prop, dataTypeFile)
    txtDatatypeFile.text = dataTypeFile
    
    Call setStetusMsg("")
    
    Call saveOdbcSettings
    
End Sub
'
' DDLを出力する
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
    
    'add tatsuo: DDLの区切り文字などの設定値
    Call Constants.clearDDLInfo(ddlInfo)
    sepStr = RTrim$(txtSepStr.text)
    commentStr = RTrim$(txtComment.text)
    ddlInfo.sepStr = sepStr
    ddlInfo.commentStr = commentStr
    
    
    '---------- Create DDL ----------
    If MsgBox("'" & ddlSheet.Name & "'のDDL定義を '" & _
                    outPath & "'へ出力します。", _
                    vbYesNo Or vbQuestion, Constants.getAppInfo) = vbYes Then
        
        Call setStetusMsg("DDL 出力中･･･")
        If erdDoc.loadTableData(ddlSheet) <> -1 Then
            Call erdDoc.writeDDL(outPath, ddlInfo)
        End If
        
    End If
    
    ' add tatsuo: DDLファイルをノートパッドで開く
    If chkDDLOpenWithNotepad Then
    
        If MsgBox("終了しました。" & vbCrLf & "メモ帳で開きますか？", vbInformation Or vbYesNo, Constants.getAppInfo) _
            = vbYes Then
            Call Util.openFileWithNotepad(outPath)
        End If
    
    Else
        ' add tatsuo:ファイルをOSに関連付けられたアプリけーションで開く。
        If MsgBox("終了しました。" & vbCrLf & "OSに関連付けれられたアプリケーションで開きますか？", vbInformation Or vbYesNo, Constants.getAppInfo) _
            = vbYes Then
             Call Util.openFileWithOS(outPath)
        End If
    
    End If
    
    Call setStetusMsg("")
    
finally:
    Call lockControls(False)
End Sub
'
' DDL設定を画面に反映
'
Private Sub setDDLSettings(newProp As Settings)
    
    Call setStetusMsg("シート位置情報設定中･･･")
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
' 数値項目をテキストにセット
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
' 書式設定を画面にセット
'
Private Sub setFormatSettings(newProp As Settings)
    
    Call setStetusMsg("書式設定中･･･")
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
' DDL出力設定を画面にセット
'
Private Sub setDDLOutputSettings(newProp As Settings)
    
    Call setStetusMsg("DDL出力設定･･･")
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
' ODBC設定を画面にセット
'
Private Sub setOdbcSettings(newProp As Settings)
    
    Call setStetusMsg("ODBC設定･･･")
    With newProp
        txtDatatypeFile.text = .getOdbcDatatypeFile
    End With
    Call setStetusMsg("")
    
End Sub
'
' UI設定を画面にセット
'
Private Sub setUISettings(newProp As Settings)
    
    ' モデルの種類の設定
    If newProp.getModelMode = 1 Then
        optLogicalModel.Value = True
    ElseIf newProp.getModelMode = 2 Then
        optPhisicaAndLogicalModel = True
    Else
        optPhisicalModel.Value = True
    End If
    
    ' 出力オプション - リレーションの出力の設定
    If newProp.getOutputRelation = 0 Then
        chkRelation.Value = False
        ' DDLの外部キー制約との連動
        chkDDLRelation.Value = False
    Else
        chkRelation.Value = True
        ' DDLの外部キー制約との連動
        chkDDLRelation.Value = True
    End If
    
    ' 出力オプション - 主キーなどの省略の設定
    If newProp.getOutputElision = 0 Then
        chkElision.Value = False
    Else
        chkElision.Value = True
    End If

    
End Sub
'
' DDL適用ボタン使用可/不可切り替え
'
Private Sub setApplicateDDLStatus(status As Boolean)
    
    cmdAppDDLSetting.Enabled = Not status

End Sub
'
' 書式適用ボタン使用可/不可切り替え
'
Private Sub setApplicateFormatStatus(status As Boolean)
    cmdAppFormatSetting.Enabled = Not status
End Sub

'
' データ定義ヘッダーをシートに挿入
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
' シートのセルに値を設定
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
' 「ｷｬﾝｾﾙ」ボタン
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
' コントロールチップを設定
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
    'add tatsuo: コメントのヒント表示
    txtComment.ControlTipText = Constants.TIPS_DDL_COMMENT

    
End Sub
'
' ユーザ設定を保存
'
Private Sub saveOdbcSettings()
    
    Call prop.setOdbcDatatypeFile(txtDatatypeFile.text)
    Call prop.saveOdbcSettings(ThisWorkbook)
        
End Sub
'
' ユーザ設定を保存
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
    
    'add tatsuo:区切り文字、コメントの文字の保存
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
' ステータスメッセージを表示
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
' DDLを開くボタンをクリックした場合
'
Private Sub openDDL_Click()
    
    Dim filePath As String
    
    If Util.isBlank(txtDDLPath.text) Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Array("出力フォルダ"))
        Exit Sub
    End If
    
    If Util.isBlank(txtDDLFile.text) Then
        Call Util.showDialog(Constants.ERR_REQUIRED_FIELD, Array("ファイル名"))
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
