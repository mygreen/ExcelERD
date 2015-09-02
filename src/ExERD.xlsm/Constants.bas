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
' ExERD�A�v���P�[�V�����ŗL�萔����сA�ŗL�֐����W���[��
'
'
'

' 2005/02/14 ver 0.5.0 ����
' 2005/02/25 ver 0.6.0 DDL�o�͂ɑΉ�
' 2005/02/25 ver 0.6.1 DDL�o�͐�w��_�C�A���O�Ƀ��b�Z�[�W�ǉ�
' 2005/02/28 ver 0.6.2 ���݂��Ȃ��ˑ��\���w�肵���ꍇ�G���[���������Ă����s����C��
' 2005/02/28 ver 0.7.0 Log�o�͂�ǉ��A�������t�@�C������ ExERD.xls.ini ���� ExERD.ini�ɕύX
' 2005/03/24 ver 0.7.1 �Ō��1�e�[�u���̂ݎw�肵�ēǂݍ��ނ��Ƃ��ł��Ȃ������s����C��
' 2005/05/20 ver 0.8.0 ERD�ADDL�쐬�O�ɁA�\����ї�̏d���`�F�b�N���s��
' 2005/05/20 ver 0.8.1 ERD�ADDL�쐬�O�ɁA�f�[�^���̊ȈՃ`�F�b�N���s��
' 2005/05/20 ver 0.8.2 ERD�ADDL�쐬����ʂ̐ݒ葮���l���f����Ȃ��s����C��
' 2005/06/16 ver 0.8.3 �\����28byte�ȏ�̏ꍇ�A�u�w�肵�����O�̃A�C�e����������܂���v�G���[�����̕s����C��
' 2005/06/29 ver 0.8.4 �O���L�[���ݒ肳��Ă��邾���ŁA�ˑ��G���e�B�e�B�ƂȂ��Ă��� �u�ˑ��v����Q�Ƃ���悤�C��
' 2005/06/29 ver 0.8.5 �ˑ��\����ݒ肷�鍀�ڂ��u�ˑ��\��.�񖼁v�ɑΉ�
' 2005/06/29 ver 0.9.0 DDL�o�͂ɂĊO���L�[����̏o�͂ɑΉ�
' 2005/06/29 ver 0.9.1 DDL�o�͂ɂ�DEFAULT�l�̏o�͂ɑΉ�
' 2005/10/12 ver 0.9.2 ���O�\���{�^���ǉ�
' 2005/10/12 ver 0.9.3 �����̓r���ŃL�����Z������@�\��ǉ�
' 2005/10/12 ver 1.0.0 ODBC�o�R�Ń��o�[�X�G���W�j�A�����O����@�\��ǉ�
' 2007/05/08 ver 1.0.1 DDL�o�͎��A��L�[���ݒ�ŃG���[�������C��

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
Public Const FIELD_NEW_SHEET                As String = "(�V�K�V�[�g)"

'DDL TITLES
Public Const START_ROW                      As String = "�J�n�s"
Public Const COL_OBJECT_TYPE                As String = "���(Table/View)"
Public Const COL_LOGICAL_TABLENAME          As String = "�\��(�_��)"
Public Const COL_PHYSICAL_TABLENAME         As String = "�\��(����)"
Public Const COL_COLID                      As String = "��No."
Public Const COL_LOGICAL_COLNAME            As String = "��(�_��)"
Public Const COL_PHYSICAL_COLNAME           As String = "��(����)"
Public Const COL_DATATYPE                   As String = "�f�[�^�^"
Public Const COL_DATALENGTH                 As String = "����"
Public Const COL_NOTNULL                    As String = "�K�{"
Public Const COL_PRIMARYKEY                 As String = "��L�["
Public Const COL_FOREIGNKEY                 As String = "�O���L�["
Public Const COL_DEPENDENCE_TABLENAME       As String = "�Q�ƕ\(��)��"
Public Const COL_RELATION_TYPE              As String = "�ˑ�"
Public Const COL_DEFAULT_VALUE              As String = "�K��l"

'DDL COMMENT
Public Const COMMENT_OBJECT_TYPE            As String = "TABLE �܂��� VIEW ���w�肵�Ă�������"
Public Const COMMENT_LOGICAL_TABLENAME      As String = "�_���e�[�u�������L�����Ă�������"
Public Const COMMENT_PHYSICAL_TABLENAME     As String = "�����e�[�u�������L�����Ă�������"
Public Const COMMENT_COLID                  As String = "��ID��A�ԂŋL�����Ă�������"
Public Const COMMENT_LOGICAL_COLNAME        As String = "�_���J���������L�����Ă�������"
Public Const COMMENT_PHYSICAL_COLNAME       As String = "�����J���������L�����Ă�������"
Public Const COMMENT_DATATYPE               As String = "�f�[�^�^���L�����Ă�������"
Public Const COMMENT_DATALENGTH             As String = "�f�[�^�^�̒���(���x)���L�����Ă�������"
Public Const COMMENT_NOTNULL                As String = "Not Null �̏ꍇ�A'Yes'���L�����Ă�������"
Public Const COMMENT_PRIMARYKEY             As String = "��L�[���ڂ̏ꍇ�A���l�܂���'Yes'���L�����Ă�������"
Public Const COMMENT_FOREIGNKEY             As String = "�O���L�[���ڂ̏ꍇ'Yes'���L�����Ă�������"
Public Const COMMENT_DEPENDENCE_TABLENAME   As String = "�O���L�[���Q�Ƃ���e�[�u�����A�������� �u�e�[�u����.�J�������v ���w�肵�Ă�������"
Public Const COMMENT_RELATION_TYPE          As String = "�Q�Ƃ���e�[�u���Ɉˑ�����ꍇ�A'Yes'���L�����Ă�������"
Public Const COMMENT_DEFAULT_VALUE          As String = "��̋K��l��ݒ肵�Ă������� ������̏ꍇ���p���ň͂񂾒l��ݒ肵�Ă������� ��- '001'"

Public Const TITLE_DDL_SEL_SHEET            As String = "�f�[�^��`�V�[�g�I��"
Public Const TITLE_ERD_SEL_SHEET            As String = "ER�}�o�̓V�[�g�I��"
Public Const TITLE_DDL_HEAD_SHEET           As String = "�f�[�^��`�V�[�g�I��(�w�b�_�[�}��)"
Public Const TITLE_MSG_FORM_PROBLEM         As String = "���̊m�F"
Public Const TITLE_REVERSE_FORM             As String = "DB���̐ݒ�"

Public Const MSG_CREATE_ERD_EXPLAIN         As String = "�e�[�u����`�������Ƃ�ER�}���쐬���܂�"
Public Const MSG_REVERSE_ERD_EXPLAIN        As String = "ODBC�o�R�Ńf�[�^�x�[�X�ɐڑ����AER�}���쐬���܂�"
Public Const MSG_CREATE_DDL_EXPLAIN         As String = "�e�[�u����`�������Ƃ�DDL���o�͂��܂�"
Public Const MSG_SHEET_POS_EXPLAIN          As String = "�e�[�u����`��񂪐ݒ肳��Ă���Excel�V�[�g�ŁA�e���ڂ��ƂɎQ�Ɨ�ʒu��ݒ肵�Ă�������"
Public Const MSG_FORMAT_EXPLAIN             As String = "ER�}�o�͏�����ݒ肵�Ă�������"
Public Const MSG_DDL_SHEET_SELECT           As String = "�e�[�u����`��񂪐ݒ肳��Ă���V�[�g��I�����Ă�������"
Public Const MSG_ERD_SHEET_SELECT           As String = "ER�}���o�͂���V�[�g��I�����Ă�������"
Public Const MSG_DDL_HEAD_SHEET_SELECT      As String = "�e�[�u����`�w�b�_�[�����o�͂���V�[�g��I�����Ă�������"
Public Const MSG_DDL_OUTPUTDIR_SELECT       As String = "DDL���o�͂���t�H���_��I�����Ă��������B"

' Control Tips
Public Const TIPS_MODEL_KIND                As String = "�o�͂���ER�}�̎�ނ�I�����Ă�������"
Public Const TIPS_RELATION                  As String = "�����[�V�������o�͂���ꍇ�`�F�b�N��ON�ɂ��Ă�������"
Public Const TIPS_DATATYPEFILE              As String = "ODBC�̃f�[�^�^��DBMS�̃f�[�^�^�̃}�b�s���O�ݒ�t�@�C�����w�肵�܂�"


Public Const TIPS_START_ROW                 As String = "�e�[�u����`��񂪐ݒ肳��Ă���擪�s"
Public Const TIPS_OBJECT_TYPE               As String = "TABLE/VIEW �敪��ݒ肷���"
Public Const TIPS_LOGICAL_TABLENAME         As String = "�_���e�[�u������ݒ肷���"
Public Const TIPS_PHISYCAL_TABLENAME        As String = "�����e�[�u������ݒ肷���"
Public Const TIPS_COLID                     As String = "�J����ID��ݒ肷���"
Public Const TIPS_LOGICAL_COLNAME           As String = "�_���J��������ݒ肷���"
Public Const TIPS_PHISYCAL_COLNAME          As String = "�����J�������ݒ肷���"
Public Const TIPS_DATATYPE                  As String = "�f�[�^�^�ݒ肷���"
Public Const TIPS_DATALENGTH                As String = "�f�[�^����(���x)��ݒ肷���"
Public Const TIPS_NOTNULL                   As String = "Not Null�����ݒ肷���"
Public Const TIPS_PRIMARYKEY                As String = "�J��������L�[(�̈ꕔ)�ł��邩��ݒ肷���"
Public Const TIPS_FOREIGNKEY                As String = "�J�������O���L�[(FK)�ł��邩��ݒ肷���"
Public Const TIPS_DEPENDENCE_TABLENAME      As String = "�O���L�[���Q�Ƃ���ꍇ�\����ݒ肷���"
Public Const TIPS_RELATION_TYPE             As String = "�O���L�[���Q�Ƃ���\�Ɉˑ����邩�ۂ���ݒ肷���"

Public Const TIPS_FONTSIZE                  As String = "ER�}�̃t�H���g�T�C�Y"
Public Const TIPS_MARGIN_LEFT               As String = "ER�}�̍��]�����w�肵�܂�"
Public Const TIPS_MARGIN_TOP                As String = "ER�}�̏�]�����w�肵�܂�"
Public Const TIPS_INTERVAL                  As String = "���f���Ԋu���w�肵�܂�"
Public Const TIPS_WIDTH_LIMIT               As String = "���f���̐܂�Ԃ��ڈ����w�肵�܂�"

Public Const TIPS_DDL_OUTPUT_DIR            As String = "DDL�o�͐�t�H���_���w�肵�Ă�������"
Public Const TIPS_DDL_OUTPUT_FILE           As String = "DDL�t�@�C�������w�肵�Ă�������"
Public Const TIPS_DDL_COMMENT               As String = "DDL�ɏo�͂���R�����g��������w�肵�Ă�������"
Public Const TIPS_DDL_SEP_TEXT              As String = "CREATE TABLE DDL�̋�؂蕶������w�肵�Ă������� Oracle�ł́A""/""�ASQLServer�ł�""GO""�Ȃ�"

'
Public Const MSG_PROBLEM_DETECT             As String = "�ȉ��̖�肪������܂����B" & vbCrLf & "�����𑱍s���܂����H"


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

Public Const REVERSE_DATATYPEFILE_FILTER    As String = "�f�[�^�^�ݒ�t�@�C��, *.dap"

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

' ERD�쐬���[�h
Public Enum ERDMODE
    Physical = &H1
    Logical = &H2
End Enum

' �R�}���h���
Public Enum CommandCondition
    CANCELL = 0&
    OK = 1&
End Enum

' �V�[�g���
Public Enum SheetMode
    DDL = 0&
    ERD = 1&
    DDL_HEAD = 2&
End Enum

'�t�@�C�����[�h
Public Enum FileMode
    AppendMode = &H1&
    BinaryMode = &H2&
    InputMode = &H3&
    OutputMode = &H4&
    Random = &H5&
End Enum

' �V�[�g���
Public Type SheetInformation
    mode        As SheetMode
    bookName    As String
    sheetName   As String
    selected    As CommandCondition
    isNewSheet  As Boolean
End Type

' ERD���
Public Type ERDInformation
    mode        As ERDMODE
    fontSize    As Single
End Type

'DDL���
Public Type DDLInformation
    sepStr      As String
    commentStr  As String
End Type

'ODBC �X�L�[�}�������
Public Type ODBCSchemaSearchParam
    catalog As Variant
    schema  As Variant
    table   As Variant
End Type

'ODBC �X�L�[�}���(�e�[�u��)
Public Type ODBCTableInfo
    tableName         As String
    tableType         As String
End Type

'ODBC �X�L�[�}���(�J����)
Public Type ODBCColumnInfo
    ordinalPosition         As String
    columnName              As String
    dataType                As String
    characterMaximumLength  As String
    numericPrecision        As String
    numericScale            As String
End Type
'
' ���b�Z�[�WID�𕶌��ɕϊ�
'
Public Function getMessage(ByVal msgId As Long) As String
    Dim result As String
    
    result = ""
    Select Case msgId
        Case ERR_NO_DDL_BOOK
            result = "�f�[�^�x�[�X��`�pExcel���[�N�u�b�N���w�肳��Ă��܂���"
        Case ERR_NO_DDL_SHEET
            result = "�f�[�^�x�[�X��`�pExcel���[�N�V�[�g���w�肳��Ă��܂���"
        Case ERR_NO_ERD_BOOK
            result = "ER�}�o�͗pExcel���[�N�u�b�N���w�肳��Ă��܂���"
        Case ERR_NO_ERD_SHEET
            result = "ER�}�o�͗pExcel���[�N�V�[�g���w�肳��Ă��܂���"
        Case ERR_NO_TABLE
            result = "�e�[�u�����w�肳��Ă��܂���"
        Case ERR_NO_DATATYPEFILE
            result = "�f�[�^�^�}�b�s���O�t�@�C���̎w�肪�s���ł��B�w�肵�Ă��������B"
        Case INFO_CANCELD_BY_USER
            result = "�L�����Z������܂���"
        Case ERR_REQUIRED_FIELD
            result = "{1}�̓��͕͂K�{�ł�"
        Case ERR_REQUIRED_MORE_VAL
            result = "{1}�ɂ�{2}�ȏ�̒l����͂��Ă�������"
        Case ERR_REQUIRED_RANGE
            result = "{1}�ɂ�{2}�`{3}�̒l����͂��Ă�������"
        Case ERR_NO_LOAD_DDL
            result = "DDL��`���ǂݍ��߂܂���ł���"
        Case ERR_ODBC_CONNECT_FAIL
            result = "ODBC �f�[�^�\�[�X�֐ڑ��ł��܂���"
        Case ERR_ODBC_NO_CONNECTION
            result = "�f�[�^�x�[�X�ɐڑ�����Ă��܂���"
        Case ERR_ODBC_NOT_SUPPORTED_OPERATION
            result = "�T�|�[�g����Ă��Ȃ���������s���܂���"
        Case ERR_ODBC_ADO_LOADING_FAIL
            result = "ActiveXDataObject�����p�ł��܂���"
        Case ERR_ODBC_GENERAL
            result = "ODBC �G���[���������܂���"
        Case ERR_GENERAL
            result = "{1}"
        Case Q_YN_CREATE_DIR
            result = "�t�H���_ [{1}] �͑��݂��܂���B�쐬���܂����H"
        Case Q_YN_OVERWRITE_FILE
            result = "[{1}] �͊��ɑ��݂��܂��B�㏑�����܂����H"
        Case Q_YN_CANCEL_PROC
            result = "�������L�����Z�����Ă�낵���ł����H"
    End Select
    
    getMessage = result
End Function
'
' �V�[�g���\���̂̏�����
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
'DDL���\���̂̏�����
'
Public Sub clearDDLInfo(ddlInfo As DDLInformation)

    With ddlInfo
        .sepStr = ""
        .commentStr = ""
    End With

End Sub
'
'ODBC �X�L�[�}�������\���̂̏�����
'
Public Sub clearODBCSchemaSearchParam(odbcParam As ODBCSchemaSearchParam)
    
    With odbcParam
        .catalog = Empty
        .schema = Empty
        .table = Empty
    End With

End Sub
'
'ODBC �X�L�[�}�������\���̂̃��Z�b�g
'
Public Sub resetODBCSchemaSearchParam(odbcParam As ODBCSchemaSearchParam)
    
    With odbcParam
        If .catalog = "" Then .catalog = Empty
        If .schema = "" Then .schema = Empty
        If .table = "" Then .table = Empty
    End With

End Sub
'
' �L�[�l�̍쐬
'   �L�[�������r�̂��߂ɓ���̕ϊ����s�Ȃ�
'   1.�g���~���O
'   2.���p
'   3.�啶��
'
Public Function keyRule(key As String)
    
    keyRule = StrConv(Trim$(key), vbNarrow Or vbUpperCase)

End Function
'
' �L�[�̃n�b�V���l���擾����
'   1.keyRule() �֐���K�p����
'   2.�L�[�̕����ɑ΂��āA'0'�` '9' �� 1 �` 10�A'A' �` 'Z' �� 11 �` 36 ��
'     �ϊ����Ēl�����v����
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
' �L�[�l�̔�r�����[���Ɋ�Â��čs�Ȃ�
'
Public Function isEqualKey(keyA As String, keyB As String) As Boolean
    isEqualKey = (keyRule(keyA) = keyRule(keyB))
End Function
'
' �V�[�g�I�𐮍����`�F�b�N
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
' APP����Ԃ�
'
Public Function getAppInfo() As String
    getAppInfo = getAppTitle() & " " & getAppVersion()
End Function
'
' APP�^�C�g����Ԃ�
'
Public Function getAppTitle() As String
    getAppTitle = APP_TITLE
End Function

'
' Version����Ԃ�
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

