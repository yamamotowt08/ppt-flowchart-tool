Attribute VB_Name = "FlowchartMaker"
Option Explicit

'============================================================
'  フローチャート自動作成マクロ（PowerPoint 用）
'
'  使い方
'   1. フローチャートを入れたいスライドを開く
'   2. マクロ「フローチャート作成」を実行
'        → スライドに黄色い入力欄が出ます
'   3. 入力欄に手順を 1 行に 1 つずつ書く
'   4. もう一度「フローチャート作成」を実行
'        → 入力欄が消えて、フローチャートが描かれます
'   5. 直したいときは、何も選ばずに「フローチャート作成」を実行
'        → 前回の文章が入った入力欄が戻るので、直して再実行
'   ※ ツール付きのファイル（.pptm / .ppam）では、リボンの
'      「フローチャート」タブのボタンから同じことができます
'
'  書き方
'   開始 / 終了 など           → 角丸（開始・終了）
'   「？」で終わる行           → ひし形（判断）
'   それ以外                   → 四角（処理）
'   判断？ > 処理A > 処理B     → 「いいえ」のとき右へ進む処理
'                                （最後の処理のあと、次の手順に合流）
'   判断？ > 処理A > #2        → 処理A のあと 2 番目の手順へ戻る
'   判断？ > #5                → 「いいえ」のとき 5 番目の手順へ
'   判断？ > 手順の文章        → 同じ文章の手順へ戻る／進む
'   判断？ > 終了              → 「いいえ」のとき右に終了
'   ※ で始まる行              → メモ（図には出ません）
'   先頭に ○ ◇ □ を付けると、形を指定できます
'============================================================

'===== 設定（お好みで変更できます）=====
Private Const YES_LABEL As String = "はい"
Private Const NO_LABEL As String = "いいえ"
Private Const START_WORDS As String = "開始|スタート|START|BEGIN|はじめ|始め|始まり"
Private Const END_WORDS As String = "終了|終わり|おわり|完了|END|STOP|エンド|ゴール"

'===== 内部で使う値 =====
Private Const INPUT_NAME As String = "フローチャート入力欄"
Private Const TAG_SRC As String = "FLOWCHART_SOURCE"
Private Const PREFIX As String = "FC_"
Private Const TITLE_TEXT As String = "フローチャート作成"

' Office の定数（参照設定がなくても動くように数値で持つ）
Private Const M_TRUE As Long = -1
Private Const M_FALSE As Long = 0
Private Const M_SHAPE_PROCESS As Long = 61
Private Const M_SHAPE_DECISION As Long = 63
Private Const M_SHAPE_TERMINATOR As Long = 69
Private Const M_CONNECTOR_STRAIGHT As Long = 1
Private Const M_ARROW_TRIANGLE As Long = 2
Private Const M_LINE_DASH As Long = 4
Private Const M_TEXT_HORIZONTAL As Long = 1
Private Const M_ALIGN_LEFT As Long = 1
Private Const M_ALIGN_CENTER As Long = 2
Private Const M_ALIGN_RIGHT As Long = 3
Private Const M_ANCHOR_TOP As Long = 1
Private Const M_ANCHOR_MIDDLE As Long = 3
Private Const M_AUTOSIZE_NONE As Long = 0
Private Const M_AUTOSIZE_SHAPE_TO_TEXT As Long = 1
Private Const M_AUTOSIZE_TEXT_TO_SHAPE As Long = 2
Private Const M_PLACEHOLDER As Long = 14

Private Const K_PROC As Long = 0
Private Const K_TERM As Long = 1
Private Const K_DEC As Long = 2

Private Const E_DOWN As Long = 1
Private Const E_RIGHT As Long = 2
Private Const E_MERGE As Long = 3
Private Const E_JUMP As Long = 4

' 図形
Private nText() As String
Private nKind() As Long
Private nEnd() As Boolean
Private nRow() As Long
Private nCol() As Long
Private nX() As Single
Private nY() As Single
Private nW() As Single
Private nH() As Single
Private nShp() As Shape
Private nCnt As Long

' 矢印
Private eFrom() As Long
Private eTo() As Long
Private eType() As Long
Private eCnt As Long

' 本線（縦に並ぶ手順）
Private mCnt As Long
Private mNode() As Long
Private mBranch() As Boolean
Private mDown() As Boolean

' 配置
Private rowHalf() As Single
Private rowPitch As Single
Private topY As Single
Private shapeNo As Long


'============================================================
'  マクロ本体
'============================================================

' 入力欄を出す → 書いたら再実行でフローチャートを描く
Public Sub フローチャート作成()
    Dim sld As Slide, inp As Shape, src As String

    Set sld = CurrentSlide()
    If sld Is Nothing Then
        MsgBox "フローチャートを入れるスライドを表示してから実行してください。", vbExclamation, TITLE_TEXT
        Exit Sub
    End If

    Set inp = SelectedTextShape()
    If Not inp Is Nothing Then
        If inp.Name <> INPUT_NAME Then
            If MsgBox("選択中のテキストを手順として、フローチャートを作ります。" & vbCr & _
                      "（作成後、このテキストは削除されます）", vbOKCancel + vbQuestion, TITLE_TEXT) <> vbOK Then Exit Sub
        End If
    Else
        Set inp = FindShapeByName(sld, INPUT_NAME)
    End If

    If inp Is Nothing Then
        src = sld.Tags(TAG_SRC)
        If Len(src) = 0 Then src = SampleText()
        CreateInputBox sld, src
        MsgBox "スライドに黄色い入力欄を出しました。" & vbCr & vbCr & _
               "手順を 1 行に 1 つずつ書いてから、" & vbCr & _
               "もう一度「フローチャート作成」を実行してください。", vbInformation, TITLE_TEXT
        Exit Sub
    End If

    src = inp.TextFrame.TextRange.Text
    If BuildFlowchart(sld, src) Then
        sld.Tags.Add TAG_SRC, src
        inp.Delete
    End If
End Sub

' 入力欄を使わず、1 行で手早く作る
Public Sub フローチャート作成_一行入力()
    Dim sld As Slide, s As String

    Set sld = CurrentSlide()
    If sld Is Nothing Then
        MsgBox "フローチャートを入れるスライドを表示してから実行してください。", vbExclamation, TITLE_TEXT
        Exit Sub
    End If

    s = InputBox("手順を「/」で区切って入力してください。" & vbCr & vbCr & _
                 "例）開始/申請書を書く/内容は正しい？ > 書き直す > #2/提出する/終了", _
                 TITLE_TEXT & "（一行入力）")
    If Len(CleanText(s)) = 0 Then Exit Sub
    s = Replace(s, "／", "/")
    s = Replace(s, "/", vbCr)
    If BuildFlowchart(sld, s) Then sld.Tags.Add TAG_SRC, s
End Sub

' 自動で描いたフローチャートを消す
Public Sub フローチャート削除()
    Dim sld As Slide

    Set sld = CurrentSlide()
    If sld Is Nothing Then Exit Sub
    If DeleteFlowShapes(sld) = 0 Then
        MsgBox "このスライドには自動作成したフローチャートがありません。", vbInformation, TITLE_TEXT
    End If
End Sub


' 書き方の説明を表示する
Public Sub フローチャートの書き方()
    MsgBox "【使い方】" & vbCr & _
           "1.［作成］→ スライドに黄色い入力欄が出ます" & vbCr & _
           "2. 手順を 1 行に 1 つずつ書いて、もう一度［作成］" & vbCr & _
           "3. 直すときは、何も選ばずに［作成］→ 入力欄が戻ります" & vbCr & vbCr & _
           "【書き方】" & vbCr & _
           "開始 / 終了　→ 角丸" & vbCr & _
           "「？」で終わる行　→ ひし形（下へ「はい」）" & vbCr & _
           "それ以外　→ 四角" & vbCr & _
           "判断？ > 処理　→「いいえ」のとき右に処理" & vbCr & _
           "判断？ > 処理 > #2　→ 処理のあと 2 番目の手順へ戻る" & vbCr & _
           "判断？ > 終了　→「いいえ」のとき右に終了" & vbCr & _
           "※ で始まる行　→ メモ（図に出ません）", vbInformation, TITLE_TEXT
End Sub


'============================================================
'  リボン「フローチャート」タブのボタン
'============================================================

Public Sub FC_RibbonMake(control As Object)
    フローチャート作成
End Sub

Public Sub FC_RibbonQuick(control As Object)
    フローチャート作成_一行入力
End Sub

Public Sub FC_RibbonDelete(control As Object)
    フローチャート削除
End Sub

Public Sub FC_RibbonHelp(control As Object)
    フローチャートの書き方
End Sub


'============================================================
'  作成の流れ
'============================================================

Private Function BuildFlowchart(sld As Slide, ByVal src As String) As Boolean
    If Not ParseSource(src) Then Exit Function
    If mCnt = 0 Then
        MsgBox "手順が 1 つも見つかりませんでした。" & vbCr & _
               "1 行に 1 つずつ手順を書いてください。", vbExclamation, TITLE_TEXT
        Exit Function
    End If

    DeleteFlowShapes sld
    DrawFlowchart sld
    BuildFlowchart = True

    If mCnt > 14 Then
        MsgBox "手順が多いため、図形が小さくなっています。" & vbCr & _
               "見づらいときは 2 枚のスライドに分けるのがおすすめです。", vbInformation, TITLE_TEXT
    End If
End Function

' 入力された文章を読み取り、図形と矢印の一覧を作る
Private Function ParseSource(ByVal src As String) As Boolean
    Dim lines() As String, lnText() As String, lnChain() As String
    Dim lnDisp() As String, lnEnd() As Boolean
    Dim items() As String, s As String, it As String, disp As String
    Dim i As Long, j As Long, p As Long, total As Long
    Dim k As Long, isEnd As Boolean
    Dim cur As Long, prv As Long, col As Long, tgt As Long, nn As Long, jumped As Boolean

    src = Replace(src, vbCrLf, vbLf)
    src = Replace(src, vbCr, vbLf)
    src = Replace(src, Chr$(11), vbLf)
    lines = Split(src & vbLf, vbLf)

    ReDim lnText(1 To UBound(lines) + 1)
    ReDim lnChain(1 To UBound(lines) + 1)
    mCnt = 0
    total = 0
    For i = 0 To UBound(lines)
        s = CleanText(lines(i))
        If Len(s) > 0 And Not IsCommentLine(s) Then
            s = Replace(s, "＞", ">")
            mCnt = mCnt + 1
            p = InStr(s, ">")
            If p > 0 Then
                lnText(mCnt) = StripListMark(CleanText(Left$(s, p - 1)))
                lnChain(mCnt) = Mid$(s, p + 1)
                total = total + UBound(Split(lnChain(mCnt), ">")) + 1
            Else
                lnText(mCnt) = StripListMark(s)
            End If
            total = total + 1
        End If
    Next i
    If mCnt = 0 Then
        ParseSource = True
        Exit Function
    End If

    ReDim nText(1 To total), nKind(1 To total), nEnd(1 To total)
    ReDim nRow(1 To total), nCol(1 To total)
    ReDim nX(1 To total), nY(1 To total), nW(1 To total), nH(1 To total)
    ReDim nShp(1 To total)
    ReDim eFrom(1 To total * 2), eTo(1 To total * 2), eType(1 To total * 2)
    ReDim mNode(1 To mCnt), mBranch(1 To mCnt), mDown(1 To mCnt)
    ReDim lnDisp(1 To mCnt), lnEnd(1 To mCnt)
    nCnt = 0
    eCnt = 0

    ' 本線の手順（縦に並ぶ）
    For i = 1 To mCnt
        disp = ClassifyStep(lnText(i), k, isEnd)
        lnDisp(i) = disp
        lnEnd(i) = isEnd
        mNode(i) = AddNodeData(disp, k, isEnd, i - 1, 0)
    Next i

    ' 矢印と「いいえ」側の処理
    For i = 1 To mCnt
        cur = mNode(i)
        If i < mCnt And Not lnEnd(i) Then
            AddEdgeData cur, mNode(i + 1), E_DOWN
            mDown(i) = True
        End If

        If Len(CleanText(lnChain(i))) > 0 Then
            items = Split(lnChain(i), ">")
            prv = cur
            col = 0
            jumped = False
            For j = 0 To UBound(items)
                it = CleanText(items(j))
                If Len(it) > 0 Then
                    tgt = FindTarget(it, lnDisp, mCnt)
                    If tgt < 0 Then
                        MsgBox i & " 番目の手順「" & lnText(i) & "」の「" & it & "」" & vbCr & vbCr & _
                               "「#番号」の番号は 1 から " & mCnt & " の間で書いてください。", vbExclamation, TITLE_TEXT
                        Exit Function
                    End If
                    If tgt > 0 Then
                        If prv <> cur And tgt = i + 1 Then
                            AddEdgeData prv, mNode(tgt), E_MERGE
                        Else
                            AddEdgeData prv, mNode(tgt), E_JUMP
                        End If
                        jumped = True
                        Exit For
                    End If
                    col = col + 1
                    disp = ClassifyStep(it, k, isEnd)
                    nn = AddNodeData(disp, k, isEnd, i - 1, col)
                    AddEdgeData prv, nn, E_RIGHT
                    prv = nn
                End If
            Next j

            ' 右側の処理が終わったら、次の手順に合流する
            If Not jumped And prv <> cur And i < mCnt Then
                If Not nEnd(prv) Then AddEdgeData prv, mNode(i + 1), E_MERGE
            End If
            mBranch(i) = jumped Or (prv <> cur)
        End If
    Next i

    ParseSource = True
End Function

Private Function AddNodeData(ByVal txt As String, ByVal kind As Long, ByVal isEnd As Boolean, _
                             ByVal r As Long, ByVal c As Long) As Long
    nCnt = nCnt + 1
    nText(nCnt) = txt
    nKind(nCnt) = kind
    nEnd(nCnt) = isEnd
    nRow(nCnt) = r
    nCol(nCnt) = c
    AddNodeData = nCnt
End Function

Private Sub AddEdgeData(ByVal a As Long, ByVal b As Long, ByVal t As Long)
    eCnt = eCnt + 1
    eFrom(eCnt) = a
    eTo(eCnt) = b
    eType(eCnt) = t
End Sub


'============================================================
'  描画
'============================================================

Private Sub DrawFlowchart(sld As Slide)
    Dim sw As Single, sh As Single
    Dim aL As Single, aR As Single, aT As Single, aB As Single
    Dim cols As Long, jumps As Long, laneSpace As Single, laneNo As Long, laneX As Single
    Dim colPitch As Single, boxW As Single, boxH As Single, decH As Single
    Dim leftX As Single, fs As Single, labelFs As Single, gy As Single
    Dim pts() As Single
    Dim i As Long, a As Long, b As Long

    sw = ActivePresentation.PageSetup.SlideWidth
    sh = ActivePresentation.PageSetup.SlideHeight
    aL = sw * 0.04
    aR = sw - sw * 0.04
    aT = ContentTop(sld, sh)
    aB = sh - sh * 0.05

    cols = 1
    For i = 1 To nCnt
        If nCol(i) + 1 > cols Then cols = nCol(i) + 1
    Next i
    jumps = 0
    For i = 1 To eCnt
        If eType(i) = E_JUMP Then jumps = jumps + 1
    Next i
    If jumps > 0 Then laneSpace = 14 + 12 * jumps

    ' 1 マスの大きさを決める
    colPitch = (aR - aL - laneSpace) / cols
    If colPitch > 240 Then colPitch = 240
    rowPitch = (aB - aT) / mCnt
    If rowPitch > 90 Then rowPitch = 90

    boxW = colPitch * 0.7
    boxH = rowPitch * 0.55
    decH = rowPitch * 0.78
    fs = boxH * 0.4
    If fs > 16 Then fs = 16
    If fs < 7 Then fs = 7
    labelFs = fs * 0.75
    If labelFs < 7 Then labelFs = 7

    ' 全体をスライドの中央に置く
    leftX = aL + ((aR - aL) - (cols * colPitch + laneSpace)) / 2
    topY = aT + ((aB - aT) - mCnt * rowPitch) / 2

    ReDim rowHalf(0 To mCnt - 1)
    For i = 1 To nCnt
        Select Case nKind(i)
            Case K_DEC
                nW(i) = boxW * 1.1
                nH(i) = decH
            Case K_TERM
                nW(i) = boxW * 0.8
                nH(i) = boxH
            Case Else
                nW(i) = boxW
                nH(i) = boxH
        End Select
        nX(i) = leftX + colPitch * (nCol(i) + 0.5)
        nY(i) = topY + rowPitch * (nRow(i) + 0.5)
        If nH(i) / 2 > rowHalf(nRow(i)) Then rowHalf(nRow(i)) = nH(i) / 2
    Next i

    shapeNo = 0
    For i = 1 To nCnt
        Set nShp(i) = AddFlowShape(sld, i, fs)
    Next i

    ' 矢印
    laneX = leftX + cols * colPitch - (colPitch - boxW) / 2 + 12
    laneNo = 0
    For i = 1 To eCnt
        a = eFrom(i)
        b = eTo(i)
        Select Case eType(i)
            Case E_DOWN
                AddConnectorLine sld, a, 3, b, 1
            Case E_RIGHT
                AddConnectorLine sld, a, 4, b, 2
            Case E_MERGE
                ' 右の処理の下から出て、次の手順の上に合流
                gy = GapAbove(b)
                ReDim pts(1 To 4, 1 To 2)
                pts(1, 1) = nX(a): pts(1, 2) = nY(a) + nH(a) / 2
                pts(2, 1) = nX(a): pts(2, 2) = gy
                pts(3, 1) = nX(b): pts(3, 2) = gy
                pts(4, 1) = nX(b): pts(4, 2) = nY(b) - nH(b) / 2
                AddPathLine sld, pts
            Case E_JUMP
                ' 右端の通り道を使って、指定の手順の上へ
                gy = GapAbove(b)
                ReDim pts(1 To 5, 1 To 2)
                pts(1, 1) = nX(a) + nW(a) / 2: pts(1, 2) = nY(a)
                pts(2, 1) = laneX + 12 * laneNo: pts(2, 2) = nY(a)
                pts(3, 1) = laneX + 12 * laneNo: pts(3, 2) = gy
                pts(4, 1) = nX(b): pts(4, 2) = gy
                pts(5, 1) = nX(b): pts(5, 2) = nY(b) - nH(b) / 2
                AddPathLine sld, pts
                laneNo = laneNo + 1
        End Select
    Next i

    ' 「はい」「いいえ」
    For i = 1 To mCnt
        a = mNode(i)
        If nKind(a) = K_DEC And mBranch(i) Then
            AddLabel sld, NO_LABEL, nX(a) + nW(a) / 2 + 3, nY(a) - labelFs * 1.4 - 1, labelFs * 4, labelFs, False
            If mDown(i) Then
                AddLabel sld, YES_LABEL, nX(a) - 3 - labelFs * 4, nY(a) + nH(a) / 2, labelFs * 4, labelFs, True
            End If
        End If
    Next i
End Sub

' 手順 b の上のすき間の高さ（合流線を通す位置）
Private Function GapAbove(ByVal b As Long) As Single
    Dim tTop As Single, r As Long, prevBottom As Single
    tTop = nY(b) - nH(b) / 2
    r = nRow(b)
    If r = 0 Then
        GapAbove = tTop - rowPitch * 0.12
    Else
        prevBottom = topY + rowPitch * (r - 0.5) + rowHalf(r - 1)
        GapAbove = (prevBottom + tTop) / 2
    End If
End Function

Private Function AddFlowShape(sld As Slide, ByVal i As Long, ByVal fs As Single) As Shape
    Dim t As Long, fillC As Long, lineC As Long, s As Shape

    Select Case nKind(i)
        Case K_TERM
            t = M_SHAPE_TERMINATOR
            fillC = RGB(226, 239, 218): lineC = RGB(84, 130, 53)
        Case K_DEC
            t = M_SHAPE_DECISION
            fillC = RGB(255, 242, 204): lineC = RGB(191, 144, 0)
        Case Else
            t = M_SHAPE_PROCESS
            fillC = RGB(222, 235, 247): lineC = RGB(46, 117, 182)
    End Select

    Set s = sld.Shapes.AddShape(t, nX(i) - nW(i) / 2, nY(i) - nH(i) / 2, nW(i), nH(i))
    NameShape s
    s.Fill.Visible = M_TRUE
    s.Fill.Solid
    s.Fill.ForeColor.RGB = fillC
    s.Line.Visible = M_TRUE
    s.Line.ForeColor.RGB = lineC
    s.Line.Weight = 1.5

    On Error Resume Next
    s.Shadow.Visible = M_FALSE
    With s.TextFrame2
        .WordWrap = M_TRUE
        .VerticalAnchor = M_ANCHOR_MIDDLE
        If nKind(i) = K_DEC Then
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        Else
            .MarginLeft = 3: .MarginRight = 3: .MarginTop = 1: .MarginBottom = 1
        End If
        .TextRange.Text = nText(i)
        .TextRange.ParagraphFormat.Alignment = M_ALIGN_CENTER
        If nKind(i) = K_DEC Then
            .TextRange.Font.Size = fs * 0.9
        Else
            .TextRange.Font.Size = fs
        End If
        .TextRange.Font.Bold = M_FALSE
        .TextRange.Font.Fill.ForeColor.RGB = RGB(38, 38, 38)
        .AutoSize = M_AUTOSIZE_TEXT_TO_SHAPE
    End With
    On Error GoTo 0

    Set AddFlowShape = s
End Function

' 図形どうしをつなぐ矢印（図形を動かしても付いてきます）
' 辺の番号: 1=上 2=左 3=下 4=右
Private Sub AddConnectorLine(sld As Slide, ByVal a As Long, ByVal sideA As Long, _
                             ByVal b As Long, ByVal sideB As Long)
    Dim c As Shape, siteA As Long, siteB As Long
    siteA = SiteOf(sld, a, sideA)
    siteB = SiteOf(sld, b, sideB)
    Set c = sld.Shapes.AddConnector(M_CONNECTOR_STRAIGHT, 0, 0, 10, 10)
    NameShape c
    c.ConnectorFormat.BeginConnect nShp(a), siteA
    c.ConnectorFormat.EndConnect nShp(b), siteB
    StyleLine c
End Sub

' 折れ線の矢印
Private Sub AddPathLine(sld As Slide, pts() As Single)
    Dim s As Shape
    Set s = sld.Shapes.AddPolyline(pts)
    NameShape s
    s.Fill.Visible = M_FALSE
    StyleLine s
End Sub

Private Sub StyleLine(s As Shape)
    With s.Line
        .Visible = M_TRUE
        .ForeColor.RGB = RGB(89, 89, 89)
        .Weight = 1.5
        .EndArrowheadStyle = M_ARROW_TRIANGLE
    End With
End Sub

Private Sub AddLabel(sld As Slide, ByVal txt As String, ByVal x As Single, ByVal y As Single, _
                     ByVal w As Single, ByVal fs As Single, ByVal alignRight As Boolean)
    Dim s As Shape
    Set s = sld.Shapes.AddTextbox(M_TEXT_HORIZONTAL, x, y, w, fs * 1.4)
    NameShape s
    On Error Resume Next
    With s.TextFrame2
        .WordWrap = M_TRUE
        .AutoSize = M_AUTOSIZE_NONE
        .VerticalAnchor = M_ANCHOR_TOP
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .TextRange.Text = txt
        .TextRange.Font.Size = fs
        .TextRange.Font.Fill.ForeColor.RGB = RGB(89, 89, 89)
        If alignRight Then
            .TextRange.ParagraphFormat.Alignment = M_ALIGN_RIGHT
        Else
            .TextRange.ParagraphFormat.Alignment = M_ALIGN_LEFT
        End If
    End With
    s.Height = fs * 1.4
    On Error GoTo 0
End Sub

Private Sub NameShape(s As Shape)
    shapeNo = shapeNo + 1
    On Error Resume Next
    s.Name = PREFIX & Format$(shapeNo, "000")
End Sub

' 図形の上下左右に一番近い接続点の番号を調べる
Private Function SiteOf(sld As Slide, ByVal i As Long, ByVal side As Long) As Long
    Dim px As Single, py As Single
    px = nX(i)
    py = nY(i)
    Select Case side
        Case 1: py = nY(i) - nH(i) / 2
        Case 2: px = nX(i) - nW(i) / 2
        Case 3: py = nY(i) + nH(i) / 2
        Case 4: px = nX(i) + nW(i) / 2
    End Select
    SiteOf = FindSite(sld, nShp(i), px, py, side)
End Function

Private Function FindSite(sld As Slide, shp As Shape, ByVal px As Single, ByVal py As Single, _
                          ByVal defSite As Long) As Long
    Dim tmp As Shape, n As Long, i As Long, bestI As Long
    Dim bx As Single, by As Single, d As Single, best As Single, worst As Single

    FindSite = defSite
    On Error Resume Next
    n = shp.ConnectionSiteCount
    If Err.Number <> 0 Or n < 1 Then Exit Function

    Set tmp = sld.Shapes.AddConnector(M_CONNECTOR_STRAIGHT, px - 10, py - 10, px, py)
    If tmp Is Nothing Then Exit Function

    best = -1
    worst = -1
    For i = 1 To n
        Err.Clear
        tmp.ConnectorFormat.BeginConnect shp, i
        If Err.Number = 0 Then
            If tmp.HorizontalFlip Then bx = tmp.Left + tmp.Width Else bx = tmp.Left
            If tmp.VerticalFlip Then by = tmp.Top + tmp.Height Else by = tmp.Top
            d = (bx - px) ^ 2 + (by - py) ^ 2
            If best < 0 Or d < best Then
                best = d
                bestI = i
            End If
            If d > worst Then worst = d
        End If
    Next i
    tmp.Delete

    ' 位置が読み取れなかったときは標準の番号（上=1 左=2 下=3 右=4）を使う
    If bestI > 0 And worst - best > 1 Then FindSite = bestI
End Function

' タイトルがあれば、その下から描く
Private Function ContentTop(sld As Slide, ByVal sh As Single) As Single
    Dim s As Shape, t As Single
    t = sh * 0.05
    For Each s In sld.Shapes
        If IsTitleShape(s) Then
            If s.HasTextFrame Then
                If s.TextFrame.HasText Then
                    If s.Top + s.Height + 6 > t Then t = s.Top + s.Height + 6
                End If
            End If
        End If
    Next s
    If t > sh * 0.4 Then t = sh * 0.05
    ContentTop = t
End Function


'============================================================
'  入力欄・スライド操作
'============================================================

Private Sub CreateInputBox(sld As Slide, ByVal src As String)
    Dim sw As Single, sh As Single, s As Shape, p As Long

    sw = ActivePresentation.PageSetup.SlideWidth
    sh = ActivePresentation.PageSetup.SlideHeight
    src = Replace(src, vbCrLf, vbCr)
    src = Replace(src, vbLf, vbCr)

    Set s = sld.Shapes.AddTextbox(M_TEXT_HORIZONTAL, sw * 0.08, sh * 0.06, sw * 0.84, sh * 0.5)
    s.Name = INPUT_NAME
    s.Fill.Visible = M_TRUE
    s.Fill.Solid
    s.Fill.ForeColor.RGB = RGB(255, 253, 231)
    s.Line.Visible = M_TRUE
    s.Line.ForeColor.RGB = RGB(237, 125, 49)
    s.Line.Weight = 2
    s.Line.DashStyle = M_LINE_DASH

    On Error Resume Next
    With s.TextFrame2
        .WordWrap = M_TRUE
        .AutoSize = M_AUTOSIZE_SHAPE_TO_TEXT
        .MarginLeft = 10: .MarginRight = 10: .MarginTop = 8: .MarginBottom = 8
        .TextRange.Text = src
        .TextRange.Font.Size = 14
        .TextRange.Font.Fill.ForeColor.RGB = RGB(38, 38, 38)
        .TextRange.ParagraphFormat.Alignment = M_ALIGN_LEFT
        ' 説明の行（※）は灰色で小さく
        For p = 1 To .TextRange.Paragraphs.Count
            If IsCommentLine(CleanText(.TextRange.Paragraphs(p).Text)) Then
                .TextRange.Paragraphs(p).Font.Size = 11
                .TextRange.Paragraphs(p).Font.Fill.ForeColor.RGB = RGB(128, 128, 128)
            End If
        Next p
    End With
    s.Select
    On Error GoTo 0
End Sub

Private Function SampleText() As String
    Dim s As String
    s = "※ 1 行に 1 つ手順を書いて、もう一度「フローチャート作成」を実行" & vbCr
    s = s & "※ 開始・終了 → 角丸、「？」で終わる行 → ひし形、その他 → 四角" & vbCr
    s = s & "※ 「判断？ > 処理」… いいえ のときの処理（右側に出ます）" & vbCr
    s = s & "※ 「#番号」… その番号の手順へ戻る／進む" & vbCr
    s = s & "※ ※ で始まる行は図に出ません（消しても大丈夫です）" & vbCr
    s = s & "開始" & vbCr
    s = s & "注文を受け付ける" & vbCr
    s = s & "会員登録済み？ > 会員登録をしてもらう" & vbCr
    s = s & "在庫はある？ > 取り寄せる > #4" & vbCr
    s = s & "商品を発送する" & vbCr
    s = s & "終了"
    SampleText = s
End Function

Private Function CurrentSlide() As Slide
    On Error Resume Next
    Set CurrentSlide = ActiveWindow.View.Slide
    If CurrentSlide Is Nothing Then Set CurrentSlide = ActiveWindow.Selection.SlideRange(1)
End Function

' 選ばれている文字入りの図形（タイトル・自動作成した図形は除く）
Private Function SelectedTextShape() As Shape
    Dim s As Shape
    On Error Resume Next
    With ActiveWindow.Selection
        If .Type <> ppSelectionShapes And .Type <> ppSelectionText Then Exit Function
        If .ShapeRange.Count <> 1 Then Exit Function
        Set s = .ShapeRange(1)
    End With
    If s Is Nothing Then Exit Function
    If Left$(s.Name, Len(PREFIX)) = PREFIX Then Exit Function
    If Not s.HasTextFrame Then Exit Function
    If Not s.TextFrame.HasText Then Exit Function
    If IsTitleShape(s) Then Exit Function
    Set SelectedTextShape = s
End Function

Private Function IsTitleShape(s As Shape) As Boolean
    Dim pt As Long
    On Error Resume Next
    If s.Type <> M_PLACEHOLDER Then Exit Function
    pt = -1
    pt = s.PlaceholderFormat.Type
    IsTitleShape = (pt = ppPlaceholderTitle Or pt = ppPlaceholderCenterTitle Or pt = ppPlaceholderVerticalTitle)
End Function

Private Function FindShapeByName(sld As Slide, ByVal nm As String) As Shape
    Dim s As Shape
    For Each s In sld.Shapes
        If s.Name = nm Then
            Set FindShapeByName = s
            Exit Function
        End If
    Next s
End Function

Private Function DeleteFlowShapes(sld As Slide) As Long
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        If Left$(sld.Shapes(i).Name, Len(PREFIX)) = PREFIX Then
            sld.Shapes(i).Delete
            DeleteFlowShapes = DeleteFlowShapes + 1
        End If
    Next i
End Function


'============================================================
'  文字の判定
'============================================================

' 形を決める（戻り値は図形に書く文字）
Private Function ClassifyStep(ByVal txt As String, ByRef kind As Long, ByRef isEnd As Boolean) As String
    Dim forced As Boolean

    kind = K_PROC
    isEnd = False
    Select Case Left$(txt, 1)
        Case "○", "◯"
            kind = K_TERM: forced = True
        Case "◇", "◆"
            kind = K_DEC: forced = True
        Case "□", "■"
            kind = K_PROC: forced = True
    End Select

    If forced Then
        txt = CleanText(Mid$(txt, 2))
    ElseIf Right$(txt, 1) = "？" Or Right$(txt, 1) = "?" Then
        kind = K_DEC
    ElseIf IsWordIn(txt, START_WORDS) Or IsWordIn(txt, END_WORDS) Then
        kind = K_TERM
    End If

    If kind = K_TERM Then isEnd = IsWordIn(txt, END_WORDS)
    ClassifyStep = txt
End Function

' 「#3」なら 3、手順と同じ文章ならその番号、どちらでもなければ 0、番号違いは -1
Private Function FindTarget(ByVal it As String, lnDisp() As String, ByVal m As Long) As Long
    Dim s As String, i As Long, n As Long, digits As Long

    If Left$(it, 1) = "#" Or Left$(it, 1) = "＃" Then
        s = CleanText(Mid$(it, 2))
        i = 1
        Do While i <= Len(s) And digits < 6
            If Not IsDigitChar(Mid$(s, i, 1)) Then Exit Do
            n = n * 10 + DigitValue(Mid$(s, i, 1))
            digits = digits + 1
            i = i + 1
        Loop
        If digits = 0 Or n < 1 Or n > m Then FindTarget = -1 Else FindTarget = n
        Exit Function
    End If

    ' 開始・終了は新しい図形として描く
    If IsWordIn(it, START_WORDS) Or IsWordIn(it, END_WORDS) Then Exit Function

    For i = 1 To m
        If lnDisp(i) = it Then
            FindTarget = i
            Exit Function
        End If
    Next i
End Function

Private Function IsWordIn(ByVal txt As String, ByVal words As String) As Boolean
    txt = UCase$(CleanText(txt))
    Do While Len(txt) > 0
        If InStr("。.!！", Right$(txt, 1)) = 0 Then Exit Do
        txt = Left$(txt, Len(txt) - 1)
    Loop
    If Len(txt) = 0 Then Exit Function
    IsWordIn = InStr(1, "|" & words & "|", "|" & txt & "|", vbBinaryCompare) > 0
End Function

Private Function IsCommentLine(ByVal s As String) As Boolean
    IsCommentLine = (Left$(s, 1) = "※" Or Left$(s, 1) = "'" Or Left$(s, 2) = "//")
End Function

' 前後の空白（全角も）を取る
Private Function CleanText(ByVal s As String) As String
    Const SP As String = " " & vbTab & "　"
    Do While Len(s) > 0
        If InStr(SP, Left$(s, 1)) = 0 Then Exit Do
        s = Mid$(s, 2)
    Loop
    Do While Len(s) > 0
        If InStr(SP, Right$(s, 1)) = 0 Then Exit Do
        s = Left$(s, Len(s) - 1)
    Loop
    CleanText = s
End Function

' 行頭の「・」「-」「1.」「２）」などを取る
Private Function StripListMark(ByVal s As String) As String
    Dim i As Long

    Select Case Left$(s, 1)
        Case "・", "-", "*"
            StripListMark = CleanText(Mid$(s, 2))
            Exit Function
    End Select

    i = 1
    Do While i <= Len(s)
        If Not IsDigitChar(Mid$(s, i, 1)) Then Exit Do
        i = i + 1
    Loop
    If i > 1 And i < Len(s) Then
        If InStr(".．)）、:：", Mid$(s, i, 1)) > 0 And Not IsDigitChar(Mid$(s, i + 1, 1)) Then
            StripListMark = CleanText(Mid$(s, i + 1))
            Exit Function
        End If
    End If
    StripListMark = s
End Function

Private Function IsDigitChar(ByVal ch As String) As Boolean
    Dim w As Long
    If Len(ch) <> 1 Then Exit Function
    w = AscW(ch) And &HFFFF&
    IsDigitChar = (w >= 48 And w <= 57) Or (w >= &HFF10& And w <= &HFF19&)
End Function

Private Function DigitValue(ByVal ch As String) As Long
    Dim w As Long
    w = AscW(ch) And &HFFFF&
    If w >= &HFF10& Then DigitValue = w - &HFF10& Else DigitValue = w - 48
End Function
