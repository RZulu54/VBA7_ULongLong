VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChessX 
   Caption         =   "ChessBrainVBA64"
   ClientHeight    =   9840.001
   ClientLeft      =   20
   ClientTop       =   230
   ClientWidth     =   16230
   OleObjectBlob   =   "frmChessX.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "frmChessX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===================================================================
'= VBAChessBrainX, a chess playing winboard engine by Roger Zuehlsdorf (Copyright 2015)
'= and is based on LarsenVb by Luca Dormio(http://xoomer.virgilio.it/ludormio/download.htm)
'=
'= VBAChessBrainX is free software: you can redistribute it and/or modify
'= it under the terms of the GNU General Public License as published by
'= the Free Software Foundation, either version 3 of the License, or
'= (at your option) any later version.
'=
'= VBAChessBrainX is distributed in the hope that it will be useful,
'= but WITHOUT ANY WARRANTY; without even the implied warranty of
'= MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'= GNU General Public License for more details.
'=
'= You should have received a copy of the GNU General Public License
'= along with this program.  If not, see <http://www.gnu.org/licenses/>.
'===================================================================

Option Explicit

#If VBA7 And Win64 Then

' GUI controls
Dim oField(1 To 64) As Control
Dim oFieldEvents(1 To 64) As clsBoardField
Dim oLabelsX(1 To 8) As Control
Dim oLabelsX2(1 To 8) As Control
Dim oLabelsY(1 To 8) As Control
Dim oLabelsY2(1 To 8) As Control
Dim oPiecePics(1 To 12) As Control
Dim oPieceCnt(1 To 6) As Control
 
Dim i As Long
Dim bThinking As Boolean

Dim Result As enumEndOfGame


Private Sub chkFlipBoard_Change()
 If chkFlipBoard.Value = True Then
   FlipBoard False
 Else
   FlipBoard True
 End If
End Sub


Private Sub cmdAttackedByWhite_Click()
  Eval
  ShowBitBoard64 AttacksForColLL(COL_WHITE) And PiecesForColLL(COL_BLACK)
End Sub

Private Sub cmdAttackedByBlack_Click()
  Eval
  ShowBitBoard64 AttacksForColLL(COL_BLACK) And PiecesForColLL(COL_WHITE)
End Sub

Private Sub cmdClearBoard_Click()
 Dim i As Integer
 For i = SQ_A1 To SQ_H8
   If Board(i) <> FRAME Then Board(i) = NO_PIECE
 Next
 ShowBoard
End Sub



Private Sub SelectPiece(PieceType As Integer)
  Dim i As Integer
  For i = 1 To 12: Me.Controls("Piece" & CStr(i)).SpecialEffect = 0: Next
  SetupPiece = PieceType
  Me.Controls("Piece" & CStr(PieceType)).SpecialEffect = 3
End Sub

Private Sub cmdEndSetup_Click()
  Dim i As Integer, WKCnt As Integer, BKCnt As Integer, bPosLegal As Boolean
  
  ' Is position legal?
  bPosLegal = True: WKCnt = 0: BKCnt = 0
  For i = SQ_A1 To SQ_H8
    Select Case Board(i)
    Case WKING: WKCnt = WKCnt + 1: If WKCnt > 1 Then bPosLegal = False: MsgBox Translate("Illegal positition: only one White King allowed!")
    Case BKING: BKCnt = BKCnt + 1: If BKCnt > 1 Then bPosLegal = False: MsgBox Translate("Illegal positition: only one Black King allowed!")
    Case WPAWN, BPAWN: If Rank(i) = 1 Or Rank(i) = 8 Then bPosLegal = False:: MsgBox Translate("Illegal positition: Pawn rank must between 2 and 7!")
    End Select
  Next
  If WKCnt = 0 Then bPosLegal = False: MsgBox Translate("Illegal positition: White King needed!")
  If BKCnt = 0 Then bPosLegal = False: MsgBox Translate("Illegal positition: Black King needed!")
  If Not bPosLegal Then Exit Sub
  
  SetupBoardMode = False
  cmdClearBoard.Visible = False
  cmdEndSetup.Visible = False
  chkWOO.Visible = False
  chkWOOO.Visible = False
  chkBOO.Visible = False
  chkBOOO.Visible = False
  lblSelectPiece.Visible = False
  cmdSetup.Visible = True
  
  ' Castling
  WhiteCastled = NO_CASTLE
  BlackCastled = NO_CASTLE
  
'  GameMovesCnt = 0
  ShowMoveList
  ShowBoard
  psLastFieldClick = "": psFieldFrom = "": psFieldTarget = "": plFieldFrom = 0: plFieldTarget = 0
End Sub


Private Sub cmdSetup_Click()
  If cmdStop.Visible Then Exit Sub ' Thinking
  SetupBoardMode = True
  cmdClearBoard.Visible = True
  cmdEndSetup.Visible = True
  lblSelectPiece.Visible = True
'  chkWOO.Visible = True: chkWOO = False
'  chkWOOO.Visible = True: chkWOOO = False
'  chkBOO.Visible = True: chkBOO = False
'  chkBOOO.Visible = True: chkBOOO = False
  cmdSetup.Visible = False
  txtIO = Translate("Select piece and click at square")
End Sub


Private Sub cmdShowBitBoard_Click()
  Dim ShowLL As LongLong, Col As enumColor
  Eval
  ShowLL = 0
  
  For Col = COL_BLACK To COL_WHITE
    If (ToggleButtonWhite And Col = COL_WHITE) Or (ToggleButtonBlack And Col = COL_BLACK) Then
      If ToggleButtonPawns Then ShowLL = ShowLL Or PiecesLL(Col, PT_PAWN)
      If ToggleButtonKnights Then ShowLL = ShowLL Or PiecesLL(Col, PT_KNIGHT)
      If ToggleButtonBishops Then ShowLL = ShowLL Or PiecesLL(Col, PT_BISHOP)
      If ToggleButtonRooks Then ShowLL = ShowLL Or PiecesLL(Col, PT_ROOK)
      If ToggleButtonQueens Then ShowLL = ShowLL Or PiecesLL(Col, PT_QUEEN)
      If ToggleButtonKings Then ShowLL = ShowLL Or PiecesLL(Col, PT_KING)
    End If
  Next Col
  
  ShowBitBoard64 ShowLL
   
End Sub

Private Sub cmdShowBitBoardAttacks_Click()
  Dim ShowLL As LongLong, Col As enumColor
  
  ShowLL = 0
  Eval
  
  For Col = COL_BLACK To COL_WHITE
    If (ToggleButtonWhite And Col = COL_WHITE) Or (ToggleButtonBlack And Col = COL_BLACK) Then
      If ToggleButtonPawnAttacks Then ShowLL = ShowLL Or AttacksForColPieceLL(Col, PT_PAWN)
      If ToggleButtonKnightAttacks Then ShowLL = ShowLL Or AttacksForColPieceLL(Col, PT_KNIGHT)
      If ToggleButtonBishopAttacks Then ShowLL = ShowLL Or AttacksForColPieceLL(Col, PT_BISHOP)
      If ToggleButtonRookAttacks Then ShowLL = ShowLL Or AttacksForColPieceLL(Col, PT_ROOK)
      If ToggleButtonQueenAttacks Then ShowLL = ShowLL Or AttacksForColPieceLL(Col, PT_QUEEN)
      If ToggleButtonKingAttacks Then ShowLL = ShowLL Or AttacksForColPieceLL(Col, PT_KING)
    End If
  Next Col
  
  ShowBitBoard64 ShowLL
 
End Sub

Private Sub cmdShowEvalMoves_Click()
  Dim Col As enumColor, BestMove As TMOVE, MoveScore As Long
  If bWhiteToMove Then Col = COL_WHITE Else Col = COL_BLACK
  Ply = 0
  txtEvalMoves.Text = EvalMoves(Col, BestMove, MoveScore)
End Sub

Private Sub cmdSwitchSideToMove_Click()
  If cmdStop.Visible = True Then Exit Sub
  bWhiteToMove = Not bWhiteToMove
  ShoCOL_WHITEToMove
End Sub


Private Sub cmdWhiteMoves_Click()
  Dim MovesLL As LongLong
  Eval
  MovesLL = GetMovesLL(COL_WHITE)
  ShowBitBoard64 MovesLL
End Sub

Private Sub cmdBlackMoves_Click()
  Dim MovesLL As LongLong
  Eval
  MovesLL = GetMovesLL(COL_BLACK)
  ShowBitBoard64 MovesLL
End Sub

Public Function GetMovesLL(inCol As enumColor) As LongLong
  ' get all target squares for Incolor for moves.
  ' > attacked squares (but not own color) AND pawn moves one or two ranks forward if squares are free
  Dim PieceMovesLL As LongLong, Us As enumColor, Them As enumColor, Pt As Long
  GetMovesLL = 0
  GetAttacksLL
  
  If inCol = COL_WHITE Then
    Us = COL_WHITE: Them = COL_BLACK
  Else
    Us = COL_BLACK: Them = COL_WHITE
  End If
  
  If Us = COL_WHITE Then
    ' try moving pawns one rank to free square
    PieceMovesLL = BoardShiftUpLL(PiecesLL(Us, PT_PAWN)) And Not BoardLL ' Shift Up = 8 * ShiftRight, white pawns never at RANK8. Special case H7 for Bit63 (sign bit)
    If PieceMovesLL <> 0 Then
      GetMovesLL = PieceMovesLL ' add moving pawns one rank
      ' Pawns at Rank 2: try two ranks
      PieceMovesLL = BoardShiftUpLL(PiecesLL(Us, PT_PAWN) And Rank2_LL) And Not BoardLL ' Shift Up = 8 * ShiftRight, white pawns never at RANK8. Special case H7 for Bit63 (sign bit)
      If PieceMovesLL <> 0 Then
        PieceMovesLL = BoardShiftUpLL(PieceMovesLL) And Not BoardLL
        If PieceMovesLL <> 0 Then GetMovesLL = GetMovesLL Or PieceMovesLL
      End If
    End If
  ElseIf Us = COL_BLACK Then
    PieceMovesLL = BoardShiftDownLL(PiecesLL(Us, PT_PAWN)) And Not BoardLL ' Shift Down = 8 * ShiftLeft, black pawns never at RANK1
    If PieceMovesLL <> 0 Then
      GetMovesLL = PieceMovesLL ' add moving pawns one rank
      ' Pawns at Rank 2: try two ranks
      PieceMovesLL = BoardShiftDownLL(PiecesLL(Us, PT_PAWN) And Rank7_LL) And Not BoardLL ' Shift Up = 8 * ShiftRight, white pawns never at RANK8. Special case H7 for Bit63 (sign bit)
      If PieceMovesLL <> 0 Then
        PieceMovesLL = BoardShiftDownLL(PieceMovesLL) And Not BoardLL
        If PieceMovesLL <> 0 Then GetMovesLL = GetMovesLL Or PieceMovesLL
      End If
    End If
 End If
 GetMovesLL = GetMovesLL Or (AttacksForColPieceLL(Us, PT_PAWN) And PiecesForColLL(Them))
 
 For Pt = PT_KNIGHT To PT_KING
   GetMovesLL = GetMovesLL Or (AttacksForColPieceLL(Us, Pt) And Not PiecesForColLL(Us)) ' all squares attacked by white but not own color
 Next

End Function


Private Sub cmdZoomMinus_Click()
  If Me.Zoom > 30 Then
    Me.Zoom = Me.Zoom - 5
    Me.Width = Me.Width * 0.95
    Me.Height = Me.Height * 0.95
  End If
End Sub

Private Sub cmdZoomPlus_Click()
  Me.Zoom = Me.Zoom + 5
  Me.Width = Me.Width * 1.05
  Me.Height = Me.Height * 1.05
End Sub







Private Sub UserForm_Initialize()
  ' GUI Start: Init
  ' Application.Workbooks.Parent.Visible = False ' Don't show EXCEL
  SetVBAPathes
  ReadColors
  CreateBoard
  LoadPiecesPics
  InitTimes
  
'  InitGame
  ShowBoard
  Me.Show
End Sub

Public Sub cmdThink_Click()
  '
  '--- Start thinking for computer move
  '
  If bThinking Or SetupBoardMode Then Exit Sub
  bThinking = True
  txtIO = ""
  
  SetTimeControl
  
 ' Result = NO_MATE
  
  If bWhiteToMove And optBlack = False Then optBlack = True
  If optWhite Then bCompIsWhite = True Else bCompIsWhite = False
  
  DoEvents
  cmdThink.Caption = Translate("Thinking") & "..."
  cmdThink.Enabled = False
  cmdStop.Visible = True
  DoEvents
  
  'SendToEngine "go"
  
  If optWhite Then bCompIsWhite = True Else bCompIsWhite = False
  
  '--- Start chess engine ----------------------
 ' StartEngine
  
  
  Dim Col As enumColor, BestMove As TMOVE, MoveScore As Long
  If bWhiteToMove Then Col = COL_WHITE Else Col = COL_BLACK
  Ply = 0
  txtEvalMoves.Text = EvalMoves(Col, BestMove, MoveScore)
 
  '--- End thinking
  If BestMove.From <> 0 Then
    CleanEpPieces
    MakeMove BestMove
    PrevLastMove = LastMove
    LastMove = BestMove
    ShowBoard
    ShowMove BestMove.From, BestMove.Target
  Else
    ShowBoard
  End If
  'ShowLastMoveAtBoard
  txtIO.Text = ""
  If BestMove.From <> 0 Then
    txtIO.Text = "Move: " & MoveText(BestMove)
  End If
  
  If BestMove.From = 0 And Not InCheck() Then
    MsgBox "DRAW!  No legal move!"
  ElseIf Abs(MoveScore) = VALUE_INFINITE Or Abs(MoveScore) = VALUE_NONE Then
    If Not InCheck() Then
      MsgBox "DRAW!  No legal move!"
    ElseIf bWhiteToMove Then
      MsgBox "MATE!  BLACK has won!"
    Else
      MsgBox "MATE!  WHITE has won!"
    End If
  End If
  
  '--- Human to move
  cmdThink.Caption = Translate("Think") & " !"
  cmdThink.Enabled = True
  cmdStop.Visible = False

  bThinking = False
  ShowMoveList
  Me.Show


End Sub



Private Sub cmdFakeInput_Click()
   '--- parse command input
'    FakeInputState = True
'    cboFakeInput.SelStart = 0
'    cboFakeInput.SelLength = Len(cboFakeInput.Text)
'    cboFakeInput.SetFocus
'    SetupBoardMode = False
'
'    If InStr(FakeInput, "setboard") > 0 Then
'      InitGame
'      txtMoveList = ""
'      Erase arGameMoves()
'      GameMovesCnt = 0
'      Result = NO_MATE
'    End If
'
'    ParseCommand FakeInput
'    ShowBoard
'
'    If bWhiteToMove Then
'      optWhite.Value = True
'    Else
'      optBlack.Value = True
'    End If
'    ShoCOL_WHITEToMove
'    psLastFieldClick = "": plFieldFrom = 0: plFieldTarget = 0
End Sub

Public Sub ShowBoard()
  Dim x As Long, y As Long, Pos As Long, Piece As Long
  
  For x = 1 To 8
    For y = 1 To 8
      Pos = x + (y - 1) * 8
      Piece = Board(SQ_A1 + x - 1 + (y - 1) * 10)
      If Piece = NO_PIECE Then
        Set oField(Pos).Picture = Nothing
      ElseIf Piece >= 1 And Piece <= 12 Then
        Set oField(Pos).Picture = oPiecePics(Piece).Picture
      End If
    Next
  Next
  ResetGUIFieldColors
  
  ' Show piece counts for white; call Eval to get counts
 ' InitEval
 ' x = Eval()
'  oPieceCnt(PieceDisplayOrder(WPAWN) + 1).Caption = CStr(PieceCnt(WPAWN) - PieceCnt(BPAWN))
'  oPieceCnt(PieceDisplayOrder(WKNIGHT) + 1).Caption = CStr(PieceCnt(WKNIGHT) - PieceCnt(BKNIGHT))
'  oPieceCnt(PieceDisplayOrder(WBISHOP) + 1).Caption = CStr(PieceCnt(WBISHOP) - PieceCnt(BBISHOP))
'  oPieceCnt(PieceDisplayOrder(WROOK) + 1).Caption = CStr(PieceCnt(WROOK) - PieceCnt(BROOK))
'  oPieceCnt(PieceDisplayOrder(WQUEEN) + 1).Caption = CStr(PieceCnt(WQUEEN) - PieceCnt(BQUEEN))
'
'  ' instead of king count show total sum
'  oPieceCnt(PieceDisplayOrder(WKING) + 1).Caption = CStr(PieceCnt(WPAWN) - PieceCnt(BPAWN) + (PieceCnt(WKNIGHT) - PieceCnt(BKNIGHT)) * 3 + _
'                                                   (PieceCnt(WBISHOP) - PieceCnt(BBISHOP)) * 3 + (PieceCnt(WROOK) - PieceCnt(BROOK)) * 5 + (PieceCnt(WQUEEN) - PieceCnt(BQUEEN)) * 9)
'
  Me.Repaint
  ShoCOL_WHITEToMove
End Sub

Private Sub CreateBoard()
 '--- Create Square Images and Labels
 Dim lFieldWidth As Long, lFrameWidth As Long
 Dim x As Long, y As Long, i As Long, bBackColorIsWhite As Boolean
 
 bBackColorIsWhite = False
 lFieldWidth = Me.fraBoard.Width \ 9 ' 8 + 1xFrame
 lFrameWidth = lFieldWidth / 2
 
 For y = 1 To 8
  '--- Label board with A - H
  Set oLabelsX(y) = Me.fraBoard.Controls.Add("Forms.Label.1", "LabelX")
  With oLabelsX(y)
    .Width = lFieldWidth: .Height = lFrameWidth: .FontSize = 12: .TextAlign = 2: .Font.Bold = True
    .Left = lFrameWidth + (y - 1) * lFieldWidth: .Top = 8 * lFieldWidth + lFrameWidth
    .BackStyle = 0: .ForeColor = &H404040: .Caption = Chr$(Asc("A") - 1 + y): .BackColor = WhiteSqCol
  End With
  
  Set oLabelsX2(y) = Me.fraBoard.Controls.Add("Forms.Label.1", "LabelX2")
  With oLabelsX2(y)
    .Width = lFieldWidth: .Height = lFrameWidth: .FontSize = 12: .TextAlign = 2: .Font.Bold = True
    .Left = lFrameWidth + (y - 1) * lFieldWidth: .Top = 2 '1 * lFieldWidth
    .BackStyle = 0: .ForeColor = &H404040: .Caption = Chr$(Asc("A") - 1 + y): .BackColor = WhiteSqCol
  End With
  
  
  '--- Label board with 1 - 8
  Set oLabelsY(y) = Me.fraBoard.Controls.Add("Forms.Label.1", "LabelY")
  With oLabelsY(y)
    .Width = lFrameWidth: .Height = lFieldWidth: .FontSize = 12: .TextAlign = 2: .Font.Bold = True
    .Left = 0: .Top = (8 - y) * lFieldWidth + lFrameWidth + lFieldWidth \ 3
    .BackStyle = 0: .ForeColor = &H404040: .Caption = CStr(y): .BackColor = WhiteSqCol
  End With
  
  Set oLabelsY2(y) = Me.fraBoard.Controls.Add("Forms.Label.1", "LabelY2")
  With oLabelsY2(y)
    .Width = lFrameWidth: .Height = lFieldWidth: .FontSize = 12: .TextAlign = 2: .Font.Bold = True
    .Left = lFrameWidth + (9 - 1) * lFieldWidth: .Top = (8 - y) * lFieldWidth + lFrameWidth + lFieldWidth \ 3
    .BackStyle = 0: .ForeColor = &H404040: .Caption = CStr(y): .BackColor = WhiteSqCol
  End With
  

  '--- set square images
  For x = 1 To 8
    i = x + (y - 1) * 8
    Set oField(i) = Me.fraBoard.Controls.Add("Forms.Image.1", "Square" & i)
    
    Set oFieldEvents(i) = New clsBoardField: oFieldEvents(i).SetBoardField oField(i) ' To catch click events
    oFieldEvents(i).Name = "Square" & i
    
    With oField(i)
      .Width = lFieldWidth: .Height = lFieldWidth: .PictureSizeMode = fmPictureSizeModeZoom
      .Left = lFrameWidth + (x - 1) * lFieldWidth:  .Top = lFrameWidth + (8 - y) * lFieldWidth
      .Tag = 20 + x + (y - 1) * 10 '--- Engine field number
      If bBackColorIsWhite Then .BackColor = WhiteSqCol Else .BackColor = BlackSqCol
      bBackColorIsWhite = Not bBackColorIsWhite
    End With
  Next x
  bBackColorIsWhite = Not bBackColorIsWhite
 Next y
End Sub

Private Sub FlipBoard(bWhiteAtBottom As Boolean)
 '--- Create Square Images and Labels
 Dim lFieldWidth As Long, lFrameWidth As Long
 Dim x As Long, y As Long, i As Long
 
 lFieldWidth = Me.fraBoard.Width \ 9 ' 8 + 1xFrame
 lFrameWidth = lFieldWidth / 2
 
 For y = 1 To 8
  '--- Label board with A - H
  With oLabelsX(y)
    If bWhiteAtBottom Then
     .Left = lFrameWidth + (y - 1) * lFieldWidth
    Else
     .Left = 8 * lFieldWidth - (lFrameWidth + (y - 1) * lFieldWidth)
    End If
  End With
  
  With oLabelsX2(y)
    If bWhiteAtBottom Then
     .Left = lFrameWidth + (y - 1) * lFieldWidth
    Else
     .Left = 8 * lFieldWidth - (lFrameWidth + (y - 1) * lFieldWidth)
    End If
  End With
  
  '--- Label board with 1 - 8
  With oLabelsY(y)
    If bWhiteAtBottom Then
     .Top = (8 - y) * lFieldWidth + lFrameWidth + lFieldWidth \ 3
    Else
     .Top = (y - 1) * lFieldWidth + lFrameWidth + lFieldWidth \ 3
    End If
  End With
  
  With oLabelsY2(y)
    If bWhiteAtBottom Then
     .Top = (8 - y) * lFieldWidth + lFrameWidth + lFieldWidth \ 3
    Else
     .Top = (y - 1) * lFieldWidth + lFrameWidth + lFieldWidth \ 3
    End If
  End With
  
  '--- set square images
  For x = 1 To 8
    i = x + (y - 1) * 8
    With oField(i)
     If bWhiteAtBottom Then
       .Left = lFrameWidth + (x - 1) * lFieldWidth:  .Top = lFrameWidth + (8 - y) * lFieldWidth
     Else
       .Left = 8 * lFieldWidth - (lFrameWidth + (x - 1) * lFieldWidth): .Top = 8 * lFieldWidth - (lFrameWidth + (8 - y) * lFieldWidth)
     End If
    End With
  Next x
 Next y
End Sub




Private Function PieceDisplayOrder(Piece As Long) As Integer
  Select Case Piece
  Case WPAWN, BPAWN: PieceDisplayOrder = 0
  Case WKNIGHT, BKNIGHT: PieceDisplayOrder = 1
  Case WBISHOP, BBISHOP: PieceDisplayOrder = 2
  Case WROOK, BROOK: PieceDisplayOrder = 3
  Case WQUEEN, BQUEEN: PieceDisplayOrder = 4
  Case WKING, BKING: PieceDisplayOrder = 5
  Case Else: PieceDisplayOrder = 0
  End Select
End Function

Private Sub cmdNewGame_Click()
  If cmdStop.Visible = True Then Exit Sub ' Thinking
  'SendToEngine "new"
  InitGame
  
  txtIO = ""
'  txtMoveList = ""
  Result = NO_MATE
  ShowBoard
End Sub


Private Sub cmdStop_Click()
  If SetupBoardMode Then Exit Sub
  bTimeExit = True
  cmdThink.Caption = Translate("Think") & " !"
  cmdThink.Enabled = True
  cmdStop.Visible = False
  bThinking = False

End Sub


Private Sub SetTimeControl()
 
End Sub

Private Sub SendToEngine(isCommand As String)
 ' ParseCommand isCommand & vbCrLf
End Sub

Private Sub TranslateForm()
'  Dim ctrl As Control, sText As String, sTextEN As String
'
'  If LangCnt = 0 Or psLanguage = "EN" Then Exit Sub
'
'  For Each ctrl In Me.Controls
'    Select Case TypeName(ctrl)
'    Case "CommandButton", "Label", "OptionButton", "CheckBox", "Frame"
'      sTextEN = ctrl.Caption
'      sText = Translate(sTextEN)
'      If sText <> sTextEN Then ctrl.Caption = sText
'    End Select
'  Next ctrl
End Sub

Private Sub cmdUndo_Click()
  Dim Move As TMOVE
  If LastMove.From <> 0 Then
    UnmakeMove LastMove
    ShowBoard
    ShowLastMoveAtBoard Move
    LastMove = PrevLastMove
    PrevLastMove.From = 0
  End If
End Sub


Private Sub fraBoard_Click()
  ' board/square clicks are handled in class clsBoardField: ImageEvents_Click
End Sub

Private Sub ShowMoveCounter()
' If GameMovesCnt <= 0 Then
'   lblMoveCnt.Caption = " "
' Else
'  lblMoveCnt.Caption = CStr(1 + GameMovesCnt \ 2) & "."
' End If
End Sub


Public Sub InitTimes()

End Sub

Public Sub ReadColors()
  Dim x As Long, s As String
  WhiteSqCol = &HC0FFFF
  BlackSqCol = &H80FF&
  BoardFrameCol = &H40C0
  fraBoard.BackColor = BoardFrameCol
End Sub


Public Sub ShowMoveList()
'  Dim i As Integer
'
'  txtMoveList = ""
'  If GameMovesCnt = 0 Then Exit Sub
'  If arGameMoves(1).Piece Mod 2 = 0 Then txtMoveList = "      "
'  For i = 1 To GameMovesCnt
'    If Len(txtMoveList) > 32000 Then txtMoveList = ""
'
'    If arGameMoves(i).Piece Mod 2 = 1 Then
'      If arGameMoves(i).From > 0 Or arGameMoves(i + 1).From > 0 Then
'        txtMoveList = txtMoveList & Left(MoveText(arGameMoves(i)) & Space(6), 6)
'      End If
'    Else
'      If arGameMoves(i).From > 0 Then txtMoveList = txtMoveList & " - " & MoveText(arGameMoves(i)) & vbCrLf
'    End If
'  Next i
'  ShowMoveCounter
'  txtMoveList.SetFocus: txtMoveList.SelStart = Len(txtMoveList): txtMoveList.SelLength = 0
'
'  DoEvents
End Sub

Private Sub Piece1_Click()
  SelectPiece 1
End Sub
Private Sub Piece2_Click()
  SelectPiece 2
End Sub
Private Sub Piece3_Click()
  SelectPiece 3
End Sub
Private Sub Piece4_Click()
  SelectPiece 4
End Sub
Private Sub Piece5_Click()
  SelectPiece 5
End Sub
Private Sub Piece6_Click()
  SelectPiece 6
End Sub
Private Sub Piece7_Click()
  SelectPiece 7
End Sub
Private Sub Piece8_Click()
  SelectPiece 8
End Sub
Private Sub Piece9_Click()
  SelectPiece 9
End Sub
Private Sub Piece10_Click()
  SelectPiece 10
End Sub
Private Sub Piece11_Click()
  SelectPiece 11
End Sub
Private Sub Piece12_Click()
  SelectPiece 12
End Sub

Private Sub LoadPiecesPics()
Dim sFile As String
Dim i As Long, lFieldWidth As Long

lFieldWidth = Me.fraPieces.Width \ 6

'--- Init piece pictures
For i = 1 To 12
   Set oPiecePics(i) = Me.Controls("Piece" & CStr(i))  ' Preloaded images
Next

End Sub

#End If
