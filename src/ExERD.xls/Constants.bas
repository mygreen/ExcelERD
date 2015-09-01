Attribute VB_Name = "Constants"
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
'
' ExERDアプリケーション固有定数および、固有関数モジュール
'
'
'

' 2005/02/14 ver 0.5.0 初版
' 2005/02/25 ver 0.6.0 DDL出力に対応
' 2005/02/25 ver 0.6.1 DDL出力先指定ダイアログにメッセージ追加
' 2005/02/28 ver 0.6.2 存在しない依存表を指定した場合エラーが発生していた不具合を修正
' 2005/02/28 ver 0.7.0 Log出力を追加、初期化ファイル名を ExERD.xls.ini から ExERD.iniに変更
' 2005/03/24 ver 0.7.1 最後の1テーブルのみ指定して読み込むことができなかった不具合を修正
' 2005/05/20 ver 0.8.0 ERD、DDL作成前に、表および列の重複チェックを行う
' 2005/05/20 ver 0.8.1 ERD、DDL作成前に、データ長の簡易チェックを行う
' 2005/05/20 ver 0.8.2 ERD、DDL作成時画面の設定属性値反映されない不具合を修正
' 2005/06/16 ver 0.8.3 表名が28byte以上の場合、「指定した名前のアイテムが見つかりません」エラー発生の不具合を修正
' 2005/06/29 ver 0.8.4 外部キーが設定されているだけで、依存エンティティとなっていた 「依存」列を参照するよう修正
' 2005/06/29 ver 0.8.5 依存表名を設定する項目を「依存表名.列名」に対応
' 2005/06/29 ver 0.9.0 DDL出力にて外部キー制約の出力に対応
' 2005/06/29 ver 0.9.1 DDL出力にてDEFAULT値の出力に対応
' 2005/10/12 ver 0.9.2 ログ表示ボタン追加
' 2005/10/12 ver 0.9.3 処理の途中でキャンセルする機能を追加
' 2005/10/12 ver 1.0.0 ODBC経由でリバースエンジニアリングする機能を追加
' 2007/05/08 ver 1.0.1 DDL出力時、主キー未設定でエラー発生を修正

'----------------------------------------
' APPLICATION INFORMATION
'----------------------------------------
Public Const APP_NAME           As String = "ExcelERD"
Public Const APP_TITLE          As String = "ExcelERD"
Public Const APP_MAJOR_VER      As Integer = 1
Public Const APP_MINOR_VER      As Integer = 0
Public Const APP_RIVISION       As Integer = 1
Public Const APP_LAST_MODEFIED  As String = "2007/05/08 12:05:28 "
Public Const APP_COPY_RIGHT     As String = "copyright(C) 2005 YAGI Hiroto All Right Reserved"
Public Const APP_AUTHOR_MAIL    As String = "piroto@a-net.email.ne.jp"

'----------------------------------------
' APPLICATION VARIABLES
'----------------------------------------
Public GLOBAL_CANCEL_FLG        As Boolean

'----------------------------------------
' APPLICATION CONSTANTS
'----------------------------------------
Public Const SEP_MARGIN As Integer = 2

Public Const MARK_PK                        As String = "(PK)"
Public Const MARK_FK                        As String = "(FK)"
Public Const FIELD_NEW_SHEET                As String = "(新規シート)"

'DDL TITLES
Public Const START_ROW                      As String = "開始行"
Public Const COL_OBJECT_TYPE                As String = "種類(Table/View)"
Public Const COL_LOGICAL_TABLENAME          As String = "表名(論理)"
Public Const COL_PHYSICAL_TABLENAME         As String = "表名(物理)"
Public Const COL_COLID                      As String = "列No."
Public Const COL_LOGICAL_COLNAME            As String = "列名(論理)"
Public Const COL_PHYSICAL_COLNAME           As String = "列名(物理)"
Public Const COL_DATATYPE                   As String = "データ型"
Public Const COL_DATALENGTH                 As String = "長さ"
Public Const COL_NOTNULL                    As String = "必須"
Public Const COL_PRIMARYKEY                 As String = "主キー"
Public Const COL_FOREIGNKEY                 As String = "外部キー"
Public Const COL_DEPENDENCE_TABLENAME       As String = "参照表(列)名"
Public Const COL_RELATION_TYPE              As String = "依存"
Public Const COL_DEFAULT_VALUE              As String = "規定値"

'DDL COMMENT
Public Const COMMENT_OBJECT_TYPE            As String = "TABLE または VIEW を指定してください"
Public Const COMMENT_LOGICAL_TABLENAME      As String = "論理テーブル名を記入してください"
Public Const COMMENT_PHYSICAL_TABLENAME     As String = "物理テーブル名を記入してください"
Public Const COMMENT_COLID                  As String = "列IDを連番で記入してください"
Public Const COMMENT_LOGICAL_COLNAME        As String = "論理カラム名を記入してください"
Public Const COMMENT_PHYSICAL_COLNAME       As String = "物理カラム名を記入してください"
Public Const COMMENT_DATATYPE               As String = "データ型を記入してください"
Public Const COMMENT_DATALENGTH             As String = "データ型の長さ(精度)を記入してください"
Public Const COMMENT_NOTNULL                As String = "Not Null の場合、'Yes'を記入してください"
Public Const COMMENT_PRIMARYKEY             As String = "主キー項目の場合、数値または'Yes'を記入してください"
Public Const COMMENT_FOREIGNKEY             As String = "外部キー項目の場合'Yes'を記入してください"
Public Const COMMENT_DEPENDENCE_TABLENAME   As String = "外部キーが参照するテーブル名、もしくは 「テーブル名.カラム名」 を指定してください"
Public Const COMMENT_RELATION_TYPE          As String = "参照するテーブルに依存する場合、'Yes'を記入してください"
Public Const COMMENT_DEFAULT_VALUE          As String = "列の規定値を設定してください 文字列の場合引用符で囲んだ値を設定してください 例- '001'"

Public Const TITLE_DDL_SEL_SHEET            As String = "データ定義シート選択"
Public Const TITLE_ERD_SEL_SHEET            As String = "ER図出力シート選択"
Public Const TITLE_DDL_HEAD_SHEET           As String = "データ定義シート選択(ヘッダー挿入)"
Public Const TITLE_MSG_FORM_PROBLEM         As String = "問題の確認"
Public Const TITLE_REVERSE_FORM             As String = "DB情報の設定"

Public Const MSG_CREATE_ERD_EXPLAIN         As String = "テーブル定義情報をもとにER図を作成します"
Public Const MSG_REVERSE_ERD_EXPLAIN        As String = "ODBC経由でデータベースに接続し、ER図を作成します"
Public Const MSG_CREATE_DDL_EXPLAIN         As String = "テーブル定義情報をもとにDDLを出力します"
Public Const MSG_SHEET_POS_EXPLAIN          As String = "テーブル定義情報が設定されているExcelシートで、各項目ごとに参照列位置を設定してください"
Public Const MSG_FORMAT_EXPLAIN             As String = "ER図出力書式を設定してください"
Public Const MSG_DDL_SHEET_SELECT           As String = "テーブル定義情報が設定されているシートを選択してください"
Public Const MSG_ERD_SHEET_SELECT           As String = "ER図を出力するシートを選択してください"
Public Const MSG_DDL_HEAD_SHEET_SELECT      As String = "テーブル定義ヘッダー情報を出力するシートを選択してください"
Public Const MSG_DDL_OUTPUTDIR_SELECT       As String = "DDLを出力するフォルダを選択してください。"

' Control Tips
Public Const TIPS_MODEL_KIND                As String = "出力するER図の種類を選択してください"
Public Const TIPS_RELATION                  As String = "リレーションを出力する場合チェックをONにしてください"
Public Const TIPS_DATATYPEFILE              As String = "ODBCのデータ型とDBMSのデータ型のマッピング設定ファイルを指定します"


Public Const TIPS_START_ROW                 As String = "テーブル定義情報が設定されている先頭行"
Public Const TIPS_OBJECT_TYPE               As String = "TABLE/VIEW 区分を設定する列"
Public Const TIPS_LOGICAL_TABLENAME         As String = "論理テーブル名を設定する列"
Public Const TIPS_PHISYCAL_TABLENAME        As String = "物理テーブル名を設定する列"
Public Const TIPS_COLID                     As String = "カラムIDを設定する列"
Public Const TIPS_LOGICAL_COLNAME           As String = "論理カラム名を設定する列"
Public Const TIPS_PHISYCAL_COLNAME          As String = "物理カラム名設定する列"
Public Const TIPS_DATATYPE                  As String = "データ型設定する列"
Public Const TIPS_DATALENGTH                As String = "データ長さ(精度)を設定する列"
Public Const TIPS_NOTNULL                   As String = "Not Null制約を設定する列"
Public Const TIPS_PRIMARYKEY                As String = "カラムが主キー(の一部)であるかを設定する列"
Public Const TIPS_FOREIGNKEY                As String = "カラムが外部キー(FK)であるかを設定する列"
Public Const TIPS_DEPENDENCE_TABLENAME      As String = "外部キーが参照する場合表名を設定する列"
Public Const TIPS_RELATION_TYPE             As String = "外部キーが参照する表に依存するか否かを設定する列"

Public Const TIPS_FONTSIZE                  As String = "ER図のフォントサイズ"
Public Const TIPS_MARGIN_LEFT               As String = "ER図の左余白を指定します"
Public Const TIPS_MARGIN_TOP                As String = "ER図の上余白を指定します"
Public Const TIPS_INTERVAL                  As String = "モデル間隔を指定します"
Public Const TIPS_WIDTH_LIMIT               As String = "モデルの折り返し目安を指定します"

Public Const TIPS_DDL_OUTPUT_DIR            As String = "DDL出力先フォルダを指定してください"
Public Const TIPS_DDL_OUTPUT_FILE           As String = "DDLファイル名を指定してください"
Public Const TIPS_DDL_COMMENT               As String = "DDLに出力するコメント文字列を指定してください"
Public Const TIPS_DDL_SEP_TEXT              As String = "CREATE TABLE DDLの区切り文字列を指定してください Oracleでは、""/""、SQLServerでは""GO""など"

'
Public Const MSG_PROBLEM_DETECT             As String = "以下の問題が見つかりました。" & vbCrLf & "処理を続行しますか？"


'ODBC Constant
Public Const ODBC_ADO_CONN_STR              As String = "ADODB.Connection"
Public Const ODBC_ADO_RECORDSET             As String = "ADODB.Recordset"
Public Const ODBC_TYPE_TABLE                As String = "TABLE"
Public Const ODBC_TYPE_VIEW                 As String = "VIEW"
'"ALIAS"
'"TABLE"
'"SYNONYM"
'"SYSTEM TABLE"
'"VIEW"
'"GLOBAL TEMPORARY"
'"LOCAL TEMPORARY"
'"SYSTEM VIEW"

Public Const REVERSE_DATATYPEFILE_FILTER    As String = "データ型設定ファイル, *.dap"

'----------------------------------------
' APPLICATION ERRORS AND INFORMATIONS
'----------------------------------------
'ERROR
Public Const ERR_NO_DDL_BOOK                    As Long = Util.ERR_MASK + &H1&
Public Const ERR_NO_DDL_SHEET                   As Long = Util.ERR_MASK + &H2&
Public Const ERR_NO_ERD_BOOK                    As Long = Util.ERR_MASK + &H3&
Public Const ERR_NO_ERD_SHEET                   As Long = Util.ERR_MASK + &H4&
Public Const ERR_REQUIRED_FIELD                 As Long = Util.ERR_MASK + &H5&
Public Const ERR_REQUIRED_MORE_VAL              As Long = Util.ERR_MASK + &H6&
Public Const ERR_REQUIRED_RANGE                 As Long = Util.ERR_MASK + &H7&
Public Const ERR_NO_LOAD_DDL                    As Long = Util.ERR_MASK + &H8&
Public Const ERR_NO_TABLE                       As Long = Util.ERR_MASK + &H9&
Public Const ERR_NO_DATATYPEFILE                As Long = Util.ERR_MASK + &H10&
'ODBC
Public Const ODBC_SUCCESS                       As Long = Util.NO_ERROR
Public Const ERR_ODBC_CONNECT_FAIL              As Long = Util.ERR_MASK + &H21&
Public Const ERR_ODBC_NO_CONNECTION             As Long = Util.ERR_MASK + &H22&
Public Const ERR_ODBC_NOT_SUPPORTED_OPERATION   As Long = Util.ERR_MASK + &H23&
Public Const ERR_ODBC_ADO_LOADING_FAIL          As Long = Util.ERR_MASK + &H24&
Public Const ERR_ODBC_GENERAL                   As Long = Util.ERR_MASK + &H2F&

Public Const ERR_GENERAL                        As Long = Util.ERR_MASK + &H3F&

'INFORMATION
Public Const INFO_CANCELD_BY_USER               As Long = Util.INFO_MASK + &H1&

'QUESTION
Public Const Q_YN_CREATE_DIR                    As Long = Util.QUESTION_YES_NO_MASK + &H1&
Public Const Q_YN_OVERWRITE_FILE                As Long = Util.QUESTION_YES_NO_MASK + &H2&
Public Const Q_YN_CANCEL_PROC                   As Long = Util.QUESTION_YES_NO_MASK + &H3&

'----------------------------------------
' APPLICATION ENUM AND TYPES
'----------------------------------------

' ERD作成モード
Public Enum ERDMODE
    Physical = &H1
    Logical = &H2
End Enum

' コマンド状態
Public Enum CommandCondition
    CANCELL = 0&
    OK = 1&
End Enum

' シート種類
Public Enum SheetMode
    DDL = 0&
    ERD = 1&
    DDL_HEAD = 2&
End Enum

'ファイルモード
Public Enum FileMode
    AppendMode = &H1&
    BinaryMode = &H2&
    InputMode = &H3&
    OutputMode = &H4&
    Random = &H5&
End Enum

' シート情報
Public Type SheetInformation
    mode        As SheetMode
    bookName    As String
    sheetName   As String
    selected    As CommandCondition
    isNewSheet  As Boolean
End Type

' ERD情報
Public Type ERDInformation
    mode        As ERDMODE
    fontSize    As Single
End Type

'DDL情報
Public Type DDLInformation
    sepStr      As String
    commentStr  As String
End Type

'ODBC スキーマ検索情報
Public Type ODBCSchemaSearchParam
    catalog As Variant
    schema  As Variant
    table   As Variant
End Type

'ODBC スキーマ情報(テーブル)
Public Type ODBCTableInfo
    tableName         As String
    tableType         As String
End Type

'ODBC スキーマ情報(カラム)
Public Type ODBCColumnInfo
    ordinalPosition         As String
    columnName              As String
    dataType                As String
    characterMaximumLength  As String
    numericPrecision        As String
    numericScale            As String
End Type
'
' メッセージIDを文言に変換
'
Public Function getMessage(ByVal msgId As Long) As String
    Dim result As String
    
    result = ""
    Select Case msgId
        Case ERR_NO_DDL_BOOK
            result = "データベース定義用Excelワークブックが指定されていません"
        Case ERR_NO_DDL_SHEET
            result = "データベース定義用Excelワークシートが指定されていません"
        Case ERR_NO_ERD_BOOK
            result = "ER図出力用Excelワークブックが指定されていません"
        Case ERR_NO_ERD_SHEET
            result = "ER図出力用Excelワークシートが指定されていません"
        Case ERR_NO_TABLE
            result = "テーブルが指定されていません"
        Case ERR_NO_DATATYPEFILE
            result = "データ型マッピングファイルの指定が不正です。指定してください。"
        Case INFO_CANCELD_BY_USER
            result = "キャンセルされました"
        Case ERR_REQUIRED_FIELD
            result = "{1}の入力は必須です"
        Case ERR_REQUIRED_MORE_VAL
            result = "{1}には{2}以上の値を入力してください"
        Case ERR_REQUIRED_RANGE
            result = "{1}には{2}～{3}の値を入力してください"
        Case ERR_NO_LOAD_DDL
            result = "DDL定義が読み込めませんでした"
        Case ERR_ODBC_CONNECT_FAIL
            result = "ODBC データソースへ接続できません"
        Case ERR_ODBC_NO_CONNECTION
            result = "データベースに接続されていません"
        Case ERR_ODBC_NOT_SUPPORTED_OPERATION
            result = "サポートされていない操作を実行しました"
        Case ERR_ODBC_ADO_LOADING_FAIL
            result = "ActiveXDataObjectが利用できません"
        Case ERR_ODBC_GENERAL
            result = "ODBC エラーが発生しました"
        Case ERR_GENERAL
            result = "{1}"
        Case Q_YN_CREATE_DIR
            result = "フォルダ [{1}] は存在しません。作成しますか？"
        Case Q_YN_OVERWRITE_FILE
            result = "[{1}] は既に存在します。上書きしますか？"
        Case Q_YN_CANCEL_PROC
            result = "処理をキャンセルしてよろしいですか？"
    End Select
    
    getMessage = result
End Function
'
' シート情報構造体の初期化
'
Public Sub clearSheetInfo(sheetInfo As SheetInformation)
    
    With sheetInfo
        .mode = DDL
        .isNewSheet = False
        .bookName = ""
        .sheetName = ""
        .selected = CommandCondition.CANCELL
    End With

End Sub
'
'DDL情報構造体の初期化
'
Public Sub clearDDLInfo(ddlInfo As DDLInformation)

    With ddlInfo
        .sepStr = ""
        .commentStr = ""
    End With

End Sub
'
'ODBC スキーマ検索情報構造体の初期化
'
Public Sub clearODBCSchemaSearchParam(odbcParam As ODBCSchemaSearchParam)
    
    With odbcParam
        .catalog = Empty
        .schema = Empty
        .table = Empty
    End With

End Sub
'
'ODBC スキーマ検索情報構造体のリセット
'
Public Sub resetODBCSchemaSearchParam(odbcParam As ODBCSchemaSearchParam)
    
    With odbcParam
        If .catalog = "" Then .catalog = Empty
        If .schema = "" Then .schema = Empty
        If .table = "" Then .table = Empty
    End With

End Sub
'
' キー値の作成
'   キー文字列比較のために同一の変換を行なう
'   1.トリミング
'   2.半角
'   3.大文字
'
Public Function keyRule(key As String)
    
    keyRule = StrConv(Trim$(key), vbNarrow Or vbUpperCase)

End Function
'
' キーのハッシュ値を取得する
'   1.keyRule() 関数を適用する
'   2.キーの文字に対して、'0'～ '9' を 1 ～ 10、'A' ～ 'Z' を 11 ～ 36 に
'     変換して値を合計する
'
Public Function getHashedKey(ByVal str As String) As String
    Dim i           As Integer
    Dim num         As Integer
    Dim c           As String
    
    Dim charBase    As Integer
    Dim numBase     As Integer
    Dim charAdjust  As Integer
    
    Dim tmpNum   As Integer
    
    numBase = Asc("0") - 1
    charBase = Asc("A") - 1
    charAdjust = Asc("9") - numBase
    
    str = keyRule(str)
    
    num = 0
    For i = 1 To Len(str)
        c = Mid$(str, i, 1)
        If IsNumeric(c) Then
            tmpNum = (Asc(c) - numBase)
        ElseIf Asc("A") <= Asc(c) And Asc(c) <= Asc("Z") Then
            tmpNum = charAdjust + (Asc(c) - charBase)
        Else
            tmpNum = (Abs(Asc(c)) Mod 99)
        End If
        num = num + tmpNum
    Next

    getHashedKey = CStr(num)
End Function
'
' キー値の比較をルールに基づいて行なう
'
Public Function isEqualKey(keyA As String, keyB As String) As Boolean
    isEqualKey = (keyRule(keyA) = keyRule(keyB))
End Function
'
' シート選択整合性チェック
'
Public Function validateSelectedSheet(sheetInfo As SheetInformation)
    
    validateSelectedSheet = Util.NO_ERROR
    
    If sheetInfo.selected <> CommandCondition.OK Then
        validateSelectedSheet = INFO_CANCELD_BY_USER
        Exit Function
    End If
    
    If (sheetInfo.mode = SheetMode.DDL) Or (sheetInfo.mode = SheetMode.DDL_HEAD) Then
        If Util.isBlank(sheetInfo.bookName) Then
            validateSelectedSheet = Constants.ERR_NO_DDL_BOOK
            Exit Function
        End If
    
        If Util.isBlank(sheetInfo.sheetName) Then
            validateSelectedSheet = Constants.ERR_NO_DDL_SHEET
            Exit Function
        End If
        
    Else
        If Util.isBlank(sheetInfo.bookName) Then
            validateSelectedSheet = Constants.ERR_NO_ERD_BOOK
            Exit Function
        End If
    
        If Util.isBlank(sheetInfo.sheetName) Then
            validateSelectedSheet = Constants.ERR_NO_ERD_SHEET
            Exit Function
        End If
    
    End If

End Function
'
' APP情報を返す
'
Public Function getAppInfo() As String
    getAppInfo = getAppTitle() & " " & getAppVersion()
End Function
'
' APPタイトルを返す
'
Public Function getAppTitle() As String
    getAppTitle = APP_TITLE
End Function

'
' Version情報を返す
'
Public Function getAppVersion() As String
    getAppVersion = APP_MAJOR_VER & "." & APP_MINOR_VER & "." & APP_RIVISION
End Function
'
'
'
Public Sub setGlobalCancelFlag(flg As Boolean)
    GLOBAL_CANCEL_FLG = flg
End Sub
'
'
'
Public Function getGlobalCancelFlag() As Boolean
    getGlobalCancelFlag = GLOBAL_CANCEL_FLG
End Function

