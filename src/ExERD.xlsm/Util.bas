Attribute VB_Name = "Util"
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
' Excel�A�v���P�[�V�����p ���� Utility�֐����W���[��
'
'

' Win32 API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpBuffer As Long) As Long

Public Const INDENT_DEPTH = 4
Public Const EXCEL_EXTENT   As String = ".xlsm"
Public Const INI_EXTENT     As String = ".ini"
Public Const LOG_EXTENT     As String = ".log"
Public Const DIR_SEP = "\"
Public Const MAX_INI_LEN As Long = 256
Public Const BIF_RETURNONLYFSDIRS = &H1      'Only return file system directories.


' ���b�Z�[�W�_�C�A���O�p�}�X�N�l
Public Const NO_ERROR               As Long = &H0&
Public Const ERR_MASK               As Long = &H40&
Public Const INFO_MASK              As Long = &H80&
Public Const QUESTION_YES_NO_MASK   As Long = &H100

'
' �t�H���_�擾�_�C�A���O��\������
'
'
' @return �I�����ꂽ�t�H���_��
' @param  title �^�C�g���̕�����
'         options    �I���I�v�V�����̒l
'         rootFolder ����t�H���_�̕�����
'
Public Function browseForFolder(title As String, _
                options, _
                Optional rootFolder As String = "") As String
    Dim cmdShell  As Object
    Dim folder As Object
    On Error GoTo errhandler
    
    Set cmdShell = CreateObject("Shell.Application")
    
    'syntax
    '  object.BrowseForFolder Hwnd, Title, Options, [RootFolder]
    '
    Set folder = cmdShell.browseForFolder(0, title, options, rootFolder)
    If Not folder Is Nothing Then
        browseForFolder = folder.Items.Item.Path
    Else
        browseForFolder = ""
    End If
    
closer:
    Set folder = Nothing
    Set cmdShell = Nothing
    
    Exit Function

errhandler:
    MsgBox Err.Description, vbCritical
    GoTo closer
    
End Function

'
'�t�H���_�擾�_�C�A���O��\������
'
Public Function browseForFolder2(title As String, _
                options, _
                Optional rootFolder As String = "") As String
    Dim cmdShell  As Object
    Dim folder As Object
    On Error GoTo errhandler
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        .InitialFileName = rootFolder
    
        If .Show = -1 Then  '�A�N�V�����{�^�����N���b�N���ꂽ
            browseForFolder2 = .SelectedItems(1)
        Else                '�L�����Z���{�^�����N���b�N���ꂽ
            browseForFolder2 = ""
        End If
    End With
    
closer:
    Set folder = Nothing
    Set cmdShell = Nothing
    
    Exit Function

errhandler:
    MsgBox Err.Description, vbCritical
    GoTo closer
    
End Function

'
'
'
Public Function chooseFile(Optional fileFilter As String, _
                            Optional rootDir As String, _
                            Optional filterIndex As Integer, _
                            Optional title As String, _
                            Optional buttonText As String, _
                            Optional MultiSelect As Boolean) As Variant
    
    If Not isBlank(Trim$(rootDir)) Then
        Call ChDrive(left$(rootDir, 1))
        Call ChDir(rootDir)
    End If
    
    chooseFile = Application.GetOpenFilename(fileFilter, filterIndex, title, buttonText, MultiSelect)
    
End Function
'
' �p�X�����쐬(�Ōオ'\'�ŏI���悤�ɐ��`����)
'
'  @return         ���`���ꂽ�p�X��
'  @param pathName ���`����p�X��
'
Public Function getPath(ByVal pathName) As String
  Dim result As String
  
  result = pathName
  If Trim$(result) = "" Then
    result = "."
  End If
  
  If Right$(result, 1) <> DIR_SEP Then
    result = result & DIR_SEP
  End If
  getPath = result
  
End Function
'
' �h���C�u���܂܂Ȃ��p�X�̏ꍇ�A�J�����g�f�B���N�g���Ƃ݂Ȃ�
'
Public Function getCurrentPath(ByVal pathName) As String
    
    If Mid$(pathName, 2, 1) <> ":" Then
        pathName = Util.getPath(CurDir$()) & pathName
    End If
    
    getCurrentPath = pathName
End Function
' 2003/4/14
' �t�@�C�����I�����ĊJ��
'
'  @return         �I�����ꂽ�t�@�C����(�L�����Z�����͋󕶎�)
'�@@param strFileFilter �t�B���^
'  @param dialogTitle �_�C�A���O�^�C�g��
'
Public Function getSelectedFilename(ByVal dialogTitle, _
                                    ByVal strFileFilter) As String
                                         
    Dim result As String
    
    '�_�C�A���O�̕\��
    result = Application.GetOpenFilename(fileFilter:=strFileFilter, _
                                            title:=dialogTitle, _
                                            MultiSelect:=False)
    If UCase$(result) = "FALSE" Then
        result = ""
    End If
    
    getSelectedFilename = result
    
End Function
' 2003/4/2
' �t�@�C�����I�����ĊJ��(�����I��)
'
'  @return         �I�����ꂽ�t�@�C����
'  @param selectedFiles �I�����ꂽ�t�@�C�������i�[���镶����z��
'�@@param strFileFilter �t�B���^
'  @param dialogTitle �_�C�A���O�^�C�g��
'
Public Function getSelectedFilenames(ByRef selectedFiles() As String, _
                                     ByVal dialogTitle, _
                                     ByVal strFileFilter) As Integer
                                         
    Dim result As Integer
    Dim fileLists As Variant
    Dim i As Integer
    Dim baseIdx As Integer
    Dim fileCount As Integer
    
    result = 0
    
    '���ʊi�[�p�z��̏�����
    ReDim selectedFiles(0)
    selectedFiles(0) = ""
    
    '�_�C�A���O�̕\��
    fileLists = Application.GetOpenFilename(fileFilter:=strFileFilter, _
                                            title:=dialogTitle, _
                                            MultiSelect:=True)
                                        
    If IsArray(fileLists) Then
        fileCount = (UBound(fileLists) - LBound(fileLists))
        baseIdx = LBound(fileLists)
        For i = 0 To fileCount
            ReDim Preserve selectedFiles(i)
            selectedFiles(i) = CStr(fileLists(baseIdx + i))
            result = result + 1
        Next
    End If
    getSelectedFilenames = result
End Function
'
' �������t�@�C�������擾����
'
Private Function getIniFilename(ByRef book As Excel.Workbook) As String
    
    getIniFilename = getWorkbookRerativeFilename(book, INI_EXTENT)

End Function
'
' ���O�t�@�C�������擾����
'
Public Function getLogFilename(ByRef book As Excel.Workbook) As String

    getLogFilename = getWorkbookRerativeFilename(book, LOG_EXTENT)
    
End Function
'
' EXCEL �t�@�C���̃v���p�X�A�g���q�ύX�������Ԃ�
'
Private Function getWorkbookRerativeFilename(ByRef book As Excel.Workbook, extents As String) As String
    Dim chk As String
    Dim tmp As String
    
    tmp = book.name
    chk = String(Len(EXCEL_EXTENT), " ") & tmp
    
    If Right$(chk, Len(EXCEL_EXTENT)) = EXCEL_EXTENT Then
        tmp = left$(tmp, Len(tmp) - Len(EXCEL_EXTENT))
    End If
    
    getWorkbookRerativeFilename = getPath(book.Path) & tmp & extents
    
End Function
'
' �ݒ��ۑ�����
'
Public Function setAppSetting(ByRef book As Excel.Workbook, _
                              ByVal strKey As String, _
                              ByVal strVal As String) As Long
    
    setAppSetting = WritePrivateProfileString( _
                        APP_NAME, strKey, strVal, getIniFilename(book))

End Function
'
' �ݒ���擾����
'
Public Function getAppSetting(ByRef book As Excel.Workbook, _
                              ByVal strKey As String, _
                              ByVal strDefault As String) As String
    
    Dim strBuf As String * MAX_INI_LEN
    Dim result As Long
    
    result = GetPrivateProfileString( _
                APP_NAME, strKey, strDefault, strBuf, MAX_INI_LEN, _
                getIniFilename(book))
                
    getAppSetting = left(strBuf, InStr(strBuf, vbNullChar) - 1)

End Function
'
' �ݒ���擾����
'
Public Function getProperty(ByVal strKey As String, _
                            ByVal strDefault As String, _
                            iniFileName As String) As String
    
    Dim strBuf As String * MAX_INI_LEN
    Dim result As Long
    
    result = GetPrivateProfileString( _
                APP_NAME, strKey, strDefault, strBuf, MAX_INI_LEN, _
                iniFileName)
                
    getProperty = left(strBuf, InStr(strBuf, vbNullChar) - 1)

End Function
' 2003/4/2
' �t�@�C�����̊g���q�𓾂�
'  @return         �g���q
'  @param filename �t�@�C����
'
Public Function getExtention(ByVal fileName As String) As String
    Dim i As Integer
    Dim c As String
    Dim result As String
    
    For i = Len(fileName) To 1 Step -1
        c = Mid$(fileName, i, 1)
        If c = "." Then
            Exit For
        Else
            result = c & result
        End If
    Next
    getExtention = result
End Function
'
' ������ɑ΂���ȈՃ\�[�g
'
Public Sub bubbleSort(ByRef values() As String)
    Dim tmp As String
    Dim i As Integer
    Dim j As Integer
    
    tmp = ""
    For i = LBound(values) To UBound(values)
        For j = i + 1 To UBound(values)
            If StrComp(values(i), values(j), vbTextCompare) > 0 Then
                tmp = values(i)
                values(i) = values(j)
                values(j) = tmp
            End If
        Next
    Next

End Sub
'
' ������z��ɁA�d�����ڂ����邩���`�F�b�N����
' �d��������ꍇ�� true ��Ԃ�
' �d�����ڂ𕶎���z��ɃZ�b�g����
'
Public Function duplicatedCheck(ByRef values() As String, ByRef duplicatedItems() As String) As Boolean
    Dim result As Boolean
    Dim i As Integer
    Dim tmp() As String
    Dim dupIdx As Integer
    
    result = False
    
    dupIdx = 0
    ReDim tmp(UBound(values))
    For i = LBound(values) To UBound(values)
        tmp(i) = values(i)
    Next
    
    Call Util.bubbleSort(tmp)
    
    Dim curVal As String
    For i = LBound(tmp) To UBound(tmp)
        If i > LBound(tmp) Then
            If curVal = tmp(i) Then
                ReDim Preserve duplicatedItems(dupIdx)
                duplicatedItems(dupIdx) = curVal
                dupIdx = dupIdx + 1
                result = True
            End If
        End If
        curVal = tmp(i)
    Next
    duplicatedCheck = result
    
End Function
'
' Proof
'
Public Sub openFileWithNotepad(fileName As String)
    On Error GoTo errhandler
    
    Call Shell("notepad " & fileName, vbNormalFocus)
    
    Exit Sub
errhandler:
    Call MsgBox(Err.Description, vbCritical, Constants.getAppInfo)
    
End Sub

'�t�@�C����OS�Ɋ֘A�t����ꂽ�A�v�����[�V�����ŊJ���B
Public Sub openFileWithOS(fileName As String)
    On Error GoTo errhandler
    
    Dim WSH
    Set WSH = CreateObject("Wscript.Shell")
    WSH.Run Chr(34) & fileName & Chr(34), 3
    Set WSH = Nothing
    
    Exit Sub
errhandler:
    Call MsgBox(Err.Description, vbCritical, Constants.getAppInfo)
    
End Sub

'
'�u�����N����
'
Public Function isBlank(str As String) As Boolean
    isBlank = ((Trim$(str)) = "")
End Function
'
'�������啶���A�������̋�ʂȂ���r����
'
Public Function compareIgnoreCase(str1 As String, str2 As String)
    compareIgnoreCase = (UCase$(str1) = UCase$(str2))
End Function
'
'����������ʂł������ĕԂ�
'
Public Function enclose(str As String) As String
    enclose = "(" & str & ")"
End Function
'
' �������Integer�ɕϊ�
'
Public Function strtoInt(str As String) As Integer
    strtoInt = CInt(val(str))
End Function
'
' Integer��Boolean�ɕϊ�
'   �l "0" �ȊO�̏ꍇ�� true �Ɣ���
'
Public Function intToBool(intVal As Integer) As Boolean
    intToBool = (Not (intVal = 0))
End Function
'
' �w�肳�ꂽ�[���̃C���f���g������(�X�y�[�X)�ŕԂ�
' 1�C���f���g�������́AINDENT_DEPTH �萔�Ŏw��
'
Public Function indent(depth As Integer) As String
    indent = Space(INDENT_DEPTH * depth)
End Function
'
' �����񂪁A�w�肳�ꂽ�����ɑ���Ȃ��ꍇ�A�X�y�[�X���߂��ĕԂ�
' �w�肳�ꂽ�����𒴂���ꍇ�́A���̂܂ܕԂ�
'
Public Function padding(str As String, length As Integer) As String
    Dim result As String
    
    If Len(str) >= length Then
        padding = str
        Exit Function
    End If

    result = str & Space(length)
    padding = left$(result, length)

End Function
'
' �w�肳�ꂽ�����񂪔z��Ɋ܂܂�邩��������
'
Public Function isContainArray(str As String, strAry() As String) As Boolean
    Dim i As Integer
    
    For i = LBound(strAry) To UBound(strAry)
        If str = strAry(i) Then
            isContainArray = True
            Exit Function
        End If
    Next
    isContainArray = False
    Exit Function
End Function

' 2014/03/01 tatsuo.tsuchie
'
' �������݉\�ȃt�H���_���ǂ�����������
' sample : http://officetanaka.net/excel/vba/tips/tips95.htm
' ���݂��Ȃ��t�H���_�̏ꍇ�Atrue��Ԃ��B
'
'
Public Function isWritableFolder(folderPath As String) As Boolean
    Dim dirStr As String
    Dim isWritable As Boolean
    Dim FSO, TmpFile As Object
    Dim tmpPath As String
    
    dirStr = Dir$(folderPath, vbDirectory)
    If dirStr = "" Then
        ' not exist folder
        isWritableFolder = True
        Exit Function
    End If
    
    ' check folder attr
    isWritable = (Not GetAttr(folderPath) And vbReadOnly)
    If Not isWritable Then
        isWritableFolder = False
        Exit Function
    End If
    
    On Error GoTo errhandler:
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    tmpPath = FSO.BuildPath(folderPath, FSO.GetTempName)
    FSO.CreateTextFile(tmpPath, True).Close
    Call Kill(tmpPath)
    
    isWritableFolder = True
    Exit Function

errhandler:
    isWritableFolder = False
    Exit Function

End Function

'
' �e��_�C�A���O�̕\�����T�|�[�g���܂�
'
'   1. ���b�Z�[�W�_�C�A���O�p�}�X�N�l + �A�� �������Ƃ��ēn�����Ƃɂ��A
'      �_�C�A���O�̕\���`���𔻒肵�܂�
'
'   2. getMessage() �֐����A�v���P�[�V�������Ƃɗp�ӂ��邱�Ƃɂ��A
'      �l �� ���b�Z�[�W�ϊ����s�Ȃ��܂�
'
'   3. ���b�Z�[�W���̒u�������� "{n}" ������ repraceStr �̒l�Œu�������܂��B
'       repraceStr ���z��̏ꍇ�A�����̒u��������ɑΉ����܂�
'
Public Function showDialog(result As Long, Optional repraceStr As Variant) As Long
    Dim btns    As Long
    Dim title   As String
    Dim msg     As String
    
    Select Case True
        Case (result And Util.ERR_MASK) = Util.ERR_MASK
            btns = vbCritical
            title = "�G���["
        Case (result And Util.INFO_MASK) = Util.INFO_MASK
            btns = vbInformation
            title = "���"
        Case (result And Util.QUESTION_YES_NO_MASK) = Util.QUESTION_YES_NO_MASK
            btns = vbQuestion Or vbYesNo
            title = "����"
        Case Else
            btns = vbOKOnly
            title = "���b�Z�[�W"
    End Select
    
    ' getMessage() �֐����A�v���P�[�V�������ƂɎ�������
    msg = Constants.getMessage(result)
    
    If Not IsMissing(repraceStr) Then
        If IsArray(repraceStr) Then
            Dim i As Integer
            For i = 0 To UBound(repraceStr)
                msg = Replace(msg, "{" & CStr(i + 1) & "}", CStr(repraceStr(i)))
            Next
        Else
            msg = Replace(msg, "{1}", CStr(repraceStr))
        End If
    End If
    
    showDialog = MsgBox(msg, vbOKOnly Or btns, getAppInfo() & " " & title)

End Function
'
'Excel version
'
Public Function getExcelVersion() As Single
    getExcelVersion = CSng(Application.version)
End Function

'
'
'
Public Sub showErrMsg(msg As String)
    
    Call Log.error(msg & " (" & Err.Number & ")" & Err.Description)
    
    Call MsgBox(msg & vbCrLf _
              & Err.Description & vbCrLf _
              & "Error No.(" & Err.Number & ")", vbCritical)
End Sub

'
'
'
Public Sub ci()
    Dim s As String
    Dim r As String
    Dim i As Integer
    Dim num As Integer
    s = InputBox("", "") 'input name full & nospace first,family
    
    r = ""
    For i = 1 To Len(s)
        num = (Asc(Mid$(s, i, 1)) Xor 255)
        r = r & CStr(num)
    Next
    s = InputBox("", "", r) 'this is password
End Sub


