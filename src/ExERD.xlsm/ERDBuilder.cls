VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ERDBuilder"
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
' ERDを作成するクラス
'
'
'

Private Const DEFAULT_FONTSIZE  As Single = 10      'デフォルトフォントサイズ

'ExcelShapeオブジェクト名プレフィックス
Private Const COL_PREFIX        As String = "C_"    'カラム
Private Const TITLE_PREFIX      As String = "L_"    'テーブル名
Private Const FRAME_PREFIX      As String = "F_"    'テーブル本体
Private Const TABLE_PREFIX      As String = "T_"    'テーブル(グループ化後)
Private Const PKLINE_PREFIX     As String = "P_"    'PKの横線
Private Const RELATION_PREFIX   As String = "R_"    '整合性制約の線

Private Const SHAPE_NAME_LIMIT  As Integer = 30     'ExcelShapeオブジェクト名の最大桁数

' メンバ変数
Private m_TargetSheet As Excel.Worksheet            '対象とするワークシート
Private m_FontSize As Single                        '設定されているフォントサイズ
'
' クラスの初期化
'
Public Sub init(ByRef sheet As Excel.Worksheet)
    Set m_TargetSheet = sheet
    Call setFontSize(DEFAULT_FONTSIZE)
End Sub

'
' このクラスが保持するシートへの参照を返す
'
Private Function getSheet() As Excel.Worksheet
    Set getSheet = m_TargetSheet
End Function
'
' 図の横幅を取得
' ・テキストボックスの幅もこの値を使用する
'
Public Function getDiagramWidth(table As ErTable, mode As ERDMODE) As Single
    Dim width As Single
    
    width = 0#
    If mode = PhysicalAndLogical Then
       
        width = width + (table.getColStringLength(Physical) * (getFontSize() * 0.7))
        width = width + (table.getColStringLength(Logical) * (getFontSize() * 0.7))
        width = width + (getFontSize() * 0.7) * 3
        width = width * 0.75
    ElseIf mode = Logical Then
        width = (table.getColStringLength(Logical) * (getFontSize() * 0.7))
    ElseIf mode = Physical Then
        width = (table.getColStringLength(Physical) * (getFontSize() * 0.7))
    Else
        width = (table.getColStringLength(Physical) * (getFontSize() * 0.7))
    End If
    
    getDiagramWidth = width
    
End Function
'
' 図の高さを取得
'
Public Function getDiagramHeight(table As ErTable) As Single
    getDiagramHeight = (table.getColCount + 3) * getCharHeight()
End Function
'
' 依存表の場合の角の丸め
'
Public Function getDiagramAdjust(table As ErTable, mode As ERDMODE) As Single
    getDiagramAdjust = getCharHeight() / (getDiagramWidth(table, mode))
End Function
'
' フォントサイズから行間の高さを得る
'
Private Function getCharHeight() As Single
    Dim version As Single
    
    version = Util.getExcelVersion()
    If version >= Constants.EXCEL2007 Then
        getCharHeight = getFontSize() * 1.7
    ElseIf version >= Constants.EXCEL2000 Then
        getCharHeight = getFontSize() * 1.4
    Else
        getCharHeight = getFontSize() * 1.4
    End If
End Function
'
' Shapeの名前の長さに制限があるため、制限を越えた場合、
' KEYの一部をハッシュ化する
'  TODO 厳密には一意性を保証できない... ほぼ大丈夫と思うが、要検討
'
Public Function createShapeName(shapeName As String) As String
    Dim nameLen As Integer
    
    nameLen = Len(shapeName)
    If nameLen <= SHAPE_NAME_LIMIT Then
        createShapeName = shapeName
        Exit Function
    End If
    
    Dim hashKey As String   'ハッシュキー
    Dim hashLen As Integer  'ハッシュキーの長さ
    Dim useLen  As Integer  '元キーの内、使用可能な桁
    
    hashKey = Constants.getHashedKey(shapeName)
    hashLen = Len(hashKey)
    
    useLen = (SHAPE_NAME_LIMIT - (hashLen + 2)) / 2
    
    createShapeName = left$(shapeName, useLen) & "_" & _
                      hashKey & "_" & _
                      Right$(shapeName, useLen)
End Function
'
' Entityオブジェクトに対して、リレーションシップを作成
'
Public Sub createRelation(parentTabele As String, childTable As String, isDepend As Boolean)
    Dim sheet As Excel.Worksheet
    Dim beginTableGroup As shape
    Dim beginShape As shape
    Dim endTableGroup As shape
    Dim endShape As shape
    
    Dim i As Integer
    
    On Error GoTo errhandler

    Set sheet = getSheet()
    sheet.Activate
    
    'Set beginShape = sheet.Shapes(TABLE_PREFIX & parentTabele).GroupItems(1)
    'Set endShape = sheet.Shapes(TABLE_PREFIX & childTable).GroupItems(1)
    Call Log.info(parentTabele & " からリレーションを開始します")
    Set beginTableGroup = sheet.Shapes(createShapeName(TABLE_PREFIX & Constants.keyRule(parentTabele)))
    
    Call Log.info(childTable & " へリレーションを接続します")
    Set endTableGroup = sheet.Shapes(createShapeName(TABLE_PREFIX & Constants.keyRule(childTable)))
    
    sheet.Shapes.AddConnector(msoConnectorElbow, 100#, 100#, 200#, 200#).Select
    Selection.Name = createShapeName(RELATION_PREFIX & parentTabele & "_" & Constants.keyRule(childTable))
    
    For i = 1 To beginTableGroup.GroupItems.Count
        Set beginShape = beginTableGroup.GroupItems(i)
        If Constants.isEqualKey(beginShape.Name, FRAME_PREFIX & parentTabele) Then
            Exit For
        End If
        Set beginShape = Nothing
    Next
    
    For i = 1 To endTableGroup.GroupItems.Count
        Set endShape = endTableGroup.GroupItems(i)
        If Constants.isEqualKey(endShape.Name, createShapeName(TITLE_PREFIX & Constants.keyRule(childTable))) Then
            Exit For
        End If
        Set endShape = Nothing
    Next
    
    If Not ((beginShape Is Nothing) Or (endShape Is Nothing)) Then
    
        Selection.ShapeRange.ConnectorFormat.BeginConnect beginShape, 3
        Selection.ShapeRange.ConnectorFormat.EndConnect endShape, 1
        
        Selection.ShapeRange.line.Weight = 0.75
        If Not isDepend Then
            Selection.ShapeRange.line.DashStyle = msoLineDash
        End If
        Selection.ShapeRange.ConnectorFormat.Type = msoConnectorElbow
        Selection.ShapeRange.line.Style = msoLineSingle
        Selection.ShapeRange.line.Transparency = 0#
        Selection.ShapeRange.line.Visible = msoTrue
        Selection.ShapeRange.line.ForeColor.SchemeColor = 64
        Selection.ShapeRange.line.BackColor.RGB = RGB(255, 255, 255)
        Selection.ShapeRange.line.BeginArrowheadLength = msoArrowheadLengthMedium
        Selection.ShapeRange.line.BeginArrowheadWidth = msoArrowheadWidthMedium
        Selection.ShapeRange.line.BeginArrowheadStyle = msoArrowheadNone
        Selection.ShapeRange.line.EndArrowheadLength = msoArrowheadShort
        Selection.ShapeRange.line.EndArrowheadWidth = msoArrowheadWidthMedium
        Selection.ShapeRange.line.EndArrowheadStyle = msoArrowheadOval
        
    End If
    
    Exit Sub
errhandler:
    
    ' 指定した名前のアイテムが見つかりませんでした。
    If Err.Number = -2147024809 Then
        Call Log.error("指定した名前のアイテムが見つかりませんでした。")
        Debug.Print Err.Description
    Else
        Call Log.error(Err.Description)
        Err.Raise Err.Number
    End If
    
End Sub
'
' Entityオブジェクトを作成
'
Public Sub createTable(ByVal mode As ERDMODE, ByRef left As Single, ByRef top As Single, ByRef tb As ErTable)
    Dim groupArry() As String
    Dim groupId   As Integer
    groupId = 0
    
    Dim titleName As String
    Dim frameName As String
    Dim pkLineName As String

    Dim sheet As Excel.Worksheet
    
    Dim width As Single
    Dim height As Single
    Dim adjust As Single
    
    Dim oldNamePrefix As String
    Dim gName As String
    Dim gShape, gItem As Object
    
    width = getDiagramWidth(tb, mode)
    height = getDiagramHeight(tb)
    
    If tb.isDependent() Or tb.isView Then
        adjust = getDiagramAdjust(tb, mode)
    Else
        adjust = 0
    End If
    
    Set sheet = getSheet()
    sheet.Activate
    
    titleName = createShapeName(TITLE_PREFIX & Constants.keyRule(tb.getPhisicName))
    frameName = createShapeName(FRAME_PREFIX & Constants.keyRule(tb.getPhisicName))
    pkLineName = createShapeName(PKLINE_PREFIX & Constants.keyRule(tb.getPhisicName))
    
    ' 重複しているオブジェクト名の変更
    ' Excel2007以降の場合、名前が重複していると、最後のグループ化に失敗するので名称を変更する。
    ' この修正により、既にRD図が存在している場合、テーブルの作成は成功するが、関連はおかしいままになるので注意
    '
    If Util.getExcelVersion >= Constants.EXCEL2007 Then
        gName = createShapeName(TABLE_PREFIX & Constants.keyRule(tb.getPhisicName))
        'random prefix
        oldNamePrefix = "_" & DateDiff("s", #1/1/1970#, Now())
        For Each gShape In ActiveSheet.Shapes
            If gShape.Name = gName Then
                For Each gItem In gShape.GroupItems
                    gItem.Name = gItem.Name & oldNamePrefix
                Next gItem
            End If
        Next gShape
    End If
    
    'Debug.Print tb.getPhisicName
    Call Log.info(tb.getPhisicName() & " を作成します")
    
    '== TITLE LABEL ==
    Call sheet.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, width, getCharHeight).Select
    
    Selection.Name = titleName
    ReDim Preserve groupArry(groupId)
    groupArry(groupId) = titleName
    groupId = groupId + 1
    
    sheet.Shapes(titleName).Select
    With Selection
        .Characters.text = tb.getTableName(mode)
        .ShapeRange.Fill.Visible = msoFalse
        .ShapeRange.Fill.Transparency = 0#
        .ShapeRange.line.Weight = 0.75
        .ShapeRange.line.DashStyle = msoLineSolid
        .ShapeRange.line.Style = msoLineSingle
        .ShapeRange.line.Transparency = 0#
        .ShapeRange.line.Visible = msoFalse
        .ShapeRange.Shadow.Visible = msoFalse
        With .Font
            .Name = "ＭＳ ゴシック"
            .size = getFontSize()
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
        End With
        
    End With
    
    If Util.getExcelVersion() >= Constants.EXCEL2007 Then
        With Selection.ShapeRange(1)
            .TextFrame2.TextRange.LanguageID = msoLanguageIDJapanese
            .TextFrame2.TextRange.Font.Name = "ＭＳ ゴシック"
            .TextFrame2.TextRange.Font.NameFarEast = "ＭＳ ゴシック"
            .TextFrame2.TextRange.Font.NameAscii = "ＭＳ ゴシック"
            .Fill.Visible = msoFalse
        End With
        
    End If
    
    top = top + getCharHeight() + getFontSize() / 2
    
    Dim colId As Integer
    Dim colLeft As Single
    Dim colHeight As Single
    Dim colTop As Single
    
    colHeight = getCharHeight()
    colTop = top + (getFontSize() / 2)
    colLeft = left + (getFontSize() / 2)
    
    ' == COLUMN INFO ==
    Call tb.sortColumn
    For colId = 0 To tb.getColCount
        Call sheet.Shapes.AddTextbox(msoTextOrientationHorizontal, colLeft, colTop, width, colHeight).Select
        Selection.Name = createShapeName(COL_PREFIX & Constants.keyRule(tb.getPhisicName & "_" & CStr(colId)))
        ReDim Preserve groupArry(groupId)
        groupArry(groupId) = Selection.Name
        groupId = groupId + 1
    
        Selection.Characters.text = tb.getColumnString(mode, colId)
        
        Selection.ShapeRange.Fill.Visible = msoFalse
        Selection.ShapeRange.Fill.Transparency = 0#
        'Selection.ShapeRange.line.Weight = 0.75
        'Selection.ShapeRange.line.DashStyle = msoLineSolid
        'Selection.ShapeRange.line.Style = msoLineSingle
        Selection.ShapeRange.line.Transparency = 0#
        Selection.ShapeRange.line.Visible = msoFalse
        Selection.ShapeRange.Shadow.Visible = msoFalse
        With Selection.Font
            .Name = "ＭＳ ゴシック"
            .size = getFontSize()
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
        End With
        
        If Util.getExcelVersion() >= Constants.EXCEL2007 Then
            With Selection.ShapeRange(1)
                .TextFrame2.TextRange.LanguageID = msoLanguageIDJapanese
                .TextFrame2.TextRange.Font.Name = "ＭＳ ゴシック"
                .TextFrame2.TextRange.Font.NameFarEast = "ＭＳ ゴシック"
                .TextFrame2.TextRange.Font.NameAscii = "ＭＳ ゴシック"
                .Fill.Visible = msoFalse
            End With
        End If

        colTop = colTop + colHeight
    Next

    '== FRAME ==
    height = colTop - top
    Call sheet.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height).Select
    
    ' for excel2007
    With Selection.ShapeRange(1)
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .line.Weight = 0.75
        .line.DashStyle = msoLineSolid
        .line.Style = msoLineSingle
        .Shadow.Visible = msoTrue
        .Shadow.Type = msoShadow14
    End With
    
    Selection.Name = frameName
    ReDim Preserve groupArry(groupId)
    groupArry(groupId) = frameName
    groupId = groupId + 1
    
    sheet.Shapes(frameName).Select
    With Selection
        .ShapeRange.ZOrder msoSendToBack
        .ShapeRange.Adjustments.Item(1) = adjust
        If tb.isView Then
            .ShapeRange.line.DashStyle = msoLineDash
        Else
            .ShapeRange.Shadow.Type = msoShadow14
        End If
    End With


    '== PRIMARY KEY SEPARATOR ==
    Dim lineTop As Single
    If (Not tb.isView) And (tb.getPKCount > 0) Then
        lineTop = top + (getFontSize() / 2) + (colHeight * tb.getPKCount)
        
        Call sheet.Shapes.AddLine(left, lineTop, left + width, lineTop).Select
        Selection.Name = pkLineName
        ReDim Preserve groupArry(groupId)
        groupArry(groupId) = Selection.Name
        groupId = groupId + 1
    End If
    
    left = left + width
    top = top + height
    
    Dim parmArry As Variant
    Dim pi As Integer
    
    groupId = groupId - 1
    parmArry = Array()
    ReDim parmArray(groupId)
    
    For pi = 0 To groupId
        parmArray(pi) = groupArry(pi)
    Next
    
    sheet.Shapes.Range(parmArray).Select
    Selection.ShapeRange.Group.Select
    Selection.Name = createShapeName(TABLE_PREFIX & Constants.keyRule(tb.getPhisicName))
End Sub

' ==== アクセサ ====
Public Sub setFontSize(fontSize As Single)
    m_FontSize = fontSize
End Sub
Public Function getFontSize() As Single
    getFontSize = m_FontSize
End Function

