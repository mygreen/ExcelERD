VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErColumun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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

'
' RDBにおけるカラムを表現するクラス
'
'
'

Private m_Pos          As Integer '位置
Private m_physicName   As String  '物理名
Private m_logicName    As String  '論理名
Private m_dataType     As String  'データ型
Private m_dataLength   As String  'データ長さ
Private m_IsPk         As Boolean 'PKか否か
Private m_IsFk         As Boolean 'FKか否か
Private m_IsNotNull    As Boolean 'NotNullか否か
Private m_RefTable     As String  '参照表名(FKの場合)
Private m_RefCol       As String  '参照列名(FKの場合)
Private m_IsDependent  As Boolean 'FKが依存か否か
Private m_defaultValue As String

'
' カラムクラスの属性を同時に設定
'
Public Sub setData(pos As Integer, logicName As String, _
                   phisicName As String, _
                   dataType As String, _
                   dataLength As String, _
                   Optional isPk As Boolean, _
                   Optional isFk As Boolean)
    
    Call setPos(pos)
    Call setLogicName(logicName)
    Call setPhysicName(phisicName)
    Call setDataType(dataType)
    Call setDataLength(dataLength)
    Call setPk(isPk)
    Call setFk(isFk)
    
End Sub
'
' このインスタンスが保持するデータが無効と判定できる場合、trueを返す
'   物理名称、論理名称、データ型のいずれも空白の場合
'
Public Function isEmpty() As Boolean
    Dim tmp As String
    
    tmp = getPhysicName()
    tmp = tmp & getLogicName()
    tmp = tmp & getDataType()
    
    isEmpty = (Len(Trim$(tmp)) = 0)

End Function
'
' ERD用のカラム文字列を指定されたモード(論理/物理)、指定された桁数に整形して返す
'
Public Function toColmunString(opt As ERDMODE, colLen As Integer) As String
    Dim colStr As String
    
    If (getColNameLength(opt) > colLen) Then
        colLen = getColNameLength(opt)
    End If
    
    If (opt = ERDMODE.Physical) Then
        colStr = getPhysicName()
    Else
        colStr = getLogicName()
    End If
    
    'PK は図で表現
    'If isPk Then
    '    colStr = colStr & ERDConstants.MARK_PK
    'End If
    
    If isFk Then
        colStr = colStr & Constants.MARK_FK
    End If
    
    
    Dim addspace As Integer
    
    addspace = colLen - getLength(colStr)
    colStr = colStr & String(addspace, " ")
    
    colStr = colStr & String(Constants.SEP_MARGIN, " ") & getDataTypeString()
    
    toColmunString = colStr
    
End Function
'
' データ型文字列表現(長さを含む)を返す
'   例) NUMBER(8,2)
'
Public Function getDataTypeString()
    Dim result As String
    
    result = getDataType()
    
    If Not Util.isBlank(getDataLength()) Then
        result = result & Util.enclose(getDataLength())
    End If

    getDataTypeString = result
    
End Function
'
' 指定されたERDモード(物理/論理)に従い、カラム名を返す
'
Public Function getColName(opt As ERDMODE) As String
    Dim result As String
    
    If opt = Logical Then
        result = getLogicName()
    Else
        result = getPhysicName
    End If

    getColName = result
End Function
'
' 指定されたERDモード(物理/論理)でのカラム文字数を返す
'
Public Function getColNameLength(opt As ERDMODE)
    Dim result As String
    
    result = 0
    If (opt = ERDMODE.Physical) Then
        result = getPhisicNameLen()
    Else
        result = getLogicNameLen()
    End If
    
    getColNameLength = result

End Function
'
' 物理名のByte長(Ascii)を取得
'
Public Function getPhisicNameLen() As Integer
    getPhisicNameLen = getLength(getPhysicName()) + getKeyMarkLen()
End Function
'
' 論理名のByte長(Ascii)を取得
'
Public Function getLogicNameLen() As Integer
    getLogicNameLen = getLength(getLogicName()) + getKeyMarkLen()
End Function
'
' データ型のByte長(Ascii)を取得
'
Public Function getDataTypeLen() As Integer
    getDataTypeLen = getLength(getDataType() & Util.enclose(getDataLength()))
End Function
'
' 文字列のByte長(Ascii)を取得
'
Private Function getLength(val As String) As Integer
    getLength = LenB(StrConv(Trim$(val), vbFromUnicode))
End Function
'
' キー表示文字列の長さを返す
'   例) このカラムがFKなら、"(FK)" の 4
'       通常のカラムの場合、0
'
Public Function getKeyMarkLen() As Integer
    Dim result As Integer
    
    result = 0
    If isFk Then
        result = getLength(Constants.MARK_FK)
    End If
    
    getKeyMarkLen = result
End Function
'
' カラムソート用のキーを返す
'   以下のルールに従う
'   1. PKの場合、PK以外のカラムより先頭
'   2. PK、PK以外のグループのなかで、getPos()の返り値の昇順
'
Public Function getSortKey() As String
    Dim prefix As String
    
    If isPk Then
        prefix = "0"
    Else
        prefix = "1"
    End If
    
    getSortKey = prefix & Format$(getPos(), "0000000000")

End Function
' ==== アクセサ ====
Public Sub setPos(pos As Integer)
    m_Pos = pos
End Sub
Public Function getPos() As Integer
    getPos = m_Pos
End Function
Public Sub setPhysicName(newVal As String)
    m_physicName = newVal
End Sub
Public Function getPhysicName() As String
    getPhysicName = m_physicName
End Function
Public Sub setLogicName(newVal As String)
    m_logicName = newVal
End Sub
Public Function getLogicName() As String
    getLogicName = m_logicName
End Function
Public Sub setDataType(newVal As String)
    m_dataType = newVal
End Sub
Public Function getDataType() As String
    getDataType = m_dataType
End Function
Public Sub setDataLength(newVal As String)
    m_dataLength = newVal
End Sub
Public Function getDataLength() As String
    getDataLength = m_dataLength
End Function
Public Sub setPk(newVal As Boolean)
    m_IsPk = newVal
End Sub
Public Function isPk() As Boolean
    isPk = m_IsPk
End Function
Public Sub setFk(newVal As Boolean)
    m_IsFk = newVal
End Sub
Public Function isFk() As Boolean
    isFk = m_IsFk
End Function
Public Sub setNotNull(newVal As Boolean)
    m_IsNotNull = newVal
End Sub
Public Function isNotNull() As Boolean
    isNotNull = m_IsNotNull
End Function
Public Function getRefTable() As String
    getRefTable = m_RefTable
End Function
Public Sub setRefTable(newVal As String)
    m_RefTable = newVal
End Sub
Public Function getRefCol() As String
    getRefCol = m_RefCol
End Function
Public Sub setRefCol(newVal As String)
    m_RefCol = newVal
End Sub
Public Sub setDependent(newVal As Boolean)
    m_IsDependent = newVal
End Sub
Public Function isDependent() As Boolean
    isDependent = m_IsDependent
End Function
Public Sub setDefaultValue(newVal As String)
    m_defaultValue = newVal
End Sub
Public Function getDefaultValue() As String
    getDefaultValue = m_defaultValue
End Function
'メンバ変数の初期化
Private Sub Class_Initialize()
    m_Pos = 0
    m_physicName = ""
    m_logicName = ""
    m_dataType = ""
    m_dataType = ""
    m_RefTable = ""
    m_RefCol = ""
    m_defaultValue = ""
    
    m_IsPk = False
    m_IsFk = False
    m_IsDependent = False

End Sub

