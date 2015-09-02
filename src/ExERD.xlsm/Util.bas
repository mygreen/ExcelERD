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
' Excelアプリケーション用 共通 Utility関数モジュール
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


' メッセージダイアログ用マスク値
Public Const NO_ERROR               As Long = &H0&
Public Const ERR_MASK               As Long = &H40&
Public Const INFO_MASK              As Long = &H80&
Public Const QUESTION_YES_NO_MASK   As Long = &H100

'
' フォルダ取得ダイアログを表示する
'
'
' @return 選択されたフォルダ名
' @param  title タイトルの文字列
'         options    選択オプションの値
'         rootFolder 既定フォルダの文字列
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
'フォルダ取得ダイアログを表示する
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
    
        If .Show = -1 Then  'アクションボタンがクリックされた
            browseForFolder2 = .SelectedItems(1)
        Else                'キャンセルボタンがクリックされた
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
' パス名を作成(最後が'\'で終わるように整形する)
'
'  @return         整形されたパス名
'  @param pathName 整形するパス名
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
' ドライブを含まないパスの場合、カレントディレクトリとみなす
'
Public Function getCurrentPath(ByVal pathName) As String
    
    If Mid$(pathName, 2, 1) <> ":" Then
        pathName = Util.getPath(CurDir$()) & pathName
    End If
    
    getCurrentPath = pathName
End Function
' 2003/4/14
' ファイル名選択して開く
'
'  @return         選択されたファイル名(キャンセル時は空文字)
'　@param strFileFilter フィルタ
'  @param dialogTitle ダイアログタイトル
'
Public Function getSelectedFilename(ByVal dialogTitle, _
                                    ByVal strFileFilter) As String
                                         
    Dim result As String
    
    'ダイアログの表示
    result = Application.GetOpenFilename(fileFilter:=strFileFilter, _
                                            title:=dialogTitle, _
                                            MultiSelect:=False)
    If UCase$(result) = "FALSE" Then
        result = ""
    End If
    
    getSelectedFilename = result
    
End Function
' 2003/4/2
' ファイル名選択して開く(複数選択)
'
'  @return         選択されたファイル数
'  @param selectedFiles 選択されたファイル名を格納する文字列配列
'　@param strFileFilter フィルタ
'  @param dialogTitle ダイアログタイトル
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
    
    '結果格納用配列の初期化
    ReDim selectedFiles(0)
    selectedFiles(0) = ""
    
    'ダイアログの表示
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
' 初期化ファイル名を取得する
'
Private Function getIniFilename(ByRef book As Excel.Workbook) As String
    
    getIniFilename = getWorkbookRerativeFilename(book, INI_EXTENT)

End Function
'
' ログファイル名を取得する
'
Public Function getLogFilename(ByRef book As Excel.Workbook) As String

    getLogFilename = getWorkbookRerativeFilename(book, LOG_EXTENT)
    
End Function
'
' EXCEL ファイルのプルパス、拡張子変更文字列を返す
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
' 設定を保存する
'
Public Function setAppSetting(ByRef book As Excel.Workbook, _
                              ByVal strKey As String, _
                              ByVal strVal As String) As Long
    
    setAppSetting = WritePrivateProfileString( _
                        APP_NAME, strKey, strVal, getIniFilename(book))

End Function
'
' 設定を取得する
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
' 設定を取得する
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
' ファイル名の拡張子を得る
'  @return         拡張子
'  @param filename ファイル名
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
' 文字列に対する簡易ソート
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
' 文字列配列に、重複項目があるかをチェックする
' 重複がある場合に true を返す
' 重複項目を文字列配列にセットする
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

'ファイルをOSに関連付けられたアプリけーションで開く。
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
'ブランク判定
'
Public Function isBlank(str As String) As Boolean
    isBlank = ((Trim$(str)) = "")
End Function
'
'文字列を大文字、小文字の区別なく比較する
'
Public Function compareIgnoreCase(str1 As String, str2 As String)
    compareIgnoreCase = (UCase$(str1) = UCase$(str2))
End Function
'
'文字列を括弧でくくって返す
'
Public Function enclose(str As String) As String
    enclose = "(" & str & ")"
End Function
'
' 文字列をIntegerに変換
'
Public Function strtoInt(str As String) As Integer
    strtoInt = CInt(val(str))
End Function
'
' IntegerをBooleanに変換
'   値 "0" 以外の場合は true と判定
'
Public Function intToBool(intVal As Integer) As Boolean
    intToBool = (Not (intVal = 0))
End Function
'
' 指定された深さのインデント文字列(スペース)で返す
' 1インデント文字数は、INDENT_DEPTH 定数で指定
'
Public Function indent(depth As Integer) As String
    indent = Space(INDENT_DEPTH * depth)
End Function
'
' 文字列が、指定された長さに足りない場合、スペース埋めして返す
' 指定された長さを超える場合は、そのまま返す
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
' 指定された文字列が配列に含まれるか検査する
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
' 書き込み可能なフォルダかどうか検査する
' sample : http://officetanaka.net/excel/vba/tips/tips95.htm
' 存在しないフォルダの場合、trueを返す。
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
' 各種ダイアログの表示をサポートします
'
'   1. メッセージダイアログ用マスク値 + 連番 を引数として渡すことにより、
'      ダイアログの表示形式を判定します
'
'   2. getMessage() 関数をアプリケーションごとに用意することにより、
'      値 → メッセージ変換を行ないます
'
'   3. メッセージ中の置換文字列 "{n}" を引数 repraceStr の値で置き換えます。
'       repraceStr が配列の場合、複数の置換文字列に対応します
'
Public Function showDialog(result As Long, Optional repraceStr As Variant) As Long
    Dim btns    As Long
    Dim title   As String
    Dim msg     As String
    
    Select Case True
        Case (result And Util.ERR_MASK) = Util.ERR_MASK
            btns = vbCritical
            title = "エラー"
        Case (result And Util.INFO_MASK) = Util.INFO_MASK
            btns = vbInformation
            title = "情報"
        Case (result And Util.QUESTION_YES_NO_MASK) = Util.QUESTION_YES_NO_MASK
            btns = vbQuestion Or vbYesNo
            title = "質問"
        Case Else
            btns = vbOKOnly
            title = "メッセージ"
    End Select
    
    ' getMessage() 関数をアプリケーションごとに実装する
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


