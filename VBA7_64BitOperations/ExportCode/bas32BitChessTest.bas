Attribute VB_Name = "bas32BitChessTest"
'=========================================================================
'= 32 bit chess test functions
'= classic chess functions with array based data structure
'= (by Roger Zuehlsdorf 2024 / email:rogzuehlsdorf@yahoo.de)
'=========================================================================
Option Explicit

#If VBA7 And Win64 Then 'Note: Win64 = Office64 bit (not Windows 64 bit)
   Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Const MAX_BOARD               As Long = 119
'--------------------------------------------------
'Piece definition
'--------------------------------------------------
'White pieces      "Board(x) mod 2 = COL_WHITE"  COL_WHITE = 1
'Black pieces      "Board(x) mod 2 = COL_BLACK"  COL_BLACK = 0
Public Const FRAME                   As Long = 0     'Frame of board array
Public Const WPAWN                   As Long = 1
Public Const BPAWN                   As Long = 2
Public Const WKNIGHT                 As Long = 3
Public Const BKNIGHT                 As Long = 4
Public Const WBISHOP                 As Long = 5
Public Const BBISHOP                 As Long = 6
Public Const WROOK                   As Long = 7
Public Const BROOK                   As Long = 8
Public Const WQUEEN                  As Long = 9
Public Const BQUEEN                  As Long = 10
Public Const WKING                   As Long = 11
Public Const BKING                   As Long = 12
Public Const NO_PIECE                As Long = 13   ' empty field
' skip 14, WEP-Piece need bit 1 set
Public Const WEP_PIECE               As Long = 15  ' en passant
Public Const BEP_PIECE               As Long = 16  ' en passant
'--- start positions
Public Const WKING_START             As Long = 25
Public Const BKING_START             As Long = 95
Public Const WQUEEN_START            As Long = 24
Public Const BQUEEN_START            As Long = 94
'--- Piece color (piece mod 2 = COL_WHITE => bit 1 set = White)
Public Const ALL_WPIECES             As Long = 14
Public Const ALL_BPIECES             As Long = 15


'--- Squares on board
Public Const SQ_A1                   As Long = 21, SQ_B1 As Long = 22, SQ_C1 As Long = 23, SQ_D1 As Long = 24, SQ_E1 As Long = 25, SQ_F1 As Long = 26, SQ_G1 As Long = 27, SQ_H1 As Long = 28
Public Const SQ_A2                   As Long = 31, SQ_B2 As Long = 32, SQ_C2 As Long = 33, SQ_D2 As Long = 34, SQ_E2 As Long = 35, SQ_F2 As Long = 36, SQ_G2 As Long = 37, SQ_H2 As Long = 38
Public Const SQ_A3                   As Long = 41, SQ_B3 As Long = 42, SQ_C3 As Long = 43, SQ_D3 As Long = 44, SQ_E3 As Long = 45, SQ_F3 As Long = 46, SQ_G3 As Long = 47, SQ_H3 As Long = 48
Public Const SQ_A4                   As Long = 51, SQ_B4 As Long = 52, SQ_C4 As Long = 53, SQ_D4 As Long = 54, SQ_E4 As Long = 55, SQ_F4 As Long = 56, SQ_G4 As Long = 57, SQ_H4 As Long = 58
Public Const SQ_A5                   As Long = 61, SQ_B5 As Long = 62, SQ_C5 As Long = 63, SQ_D5 As Long = 64, SQ_E5 As Long = 65, SQ_F5 As Long = 66, SQ_G5 As Long = 67, SQ_H5 As Long = 68
Public Const SQ_A6                   As Long = 71, SQ_B6 As Long = 72, SQ_C6 As Long = 73, SQ_D6 As Long = 74, SQ_E6 As Long = 75, SQ_F6 As Long = 76, SQ_G6 As Long = 77, SQ_H6 As Long = 78
Public Const SQ_A7                   As Long = 81, SQ_B7 As Long = 82, SQ_C7 As Long = 83, SQ_D7 As Long = 84, SQ_E7 As Long = 85, SQ_F7 As Long = 86, SQ_G7 As Long = 87, SQ_H7 As Long = 88
Public Const SQ_A8                   As Long = 91, SQ_B8 As Long = 92, SQ_C8 As Long = 93, SQ_D8 As Long = 94, SQ_E8 As Long = 95, SQ_F8 As Long = 96, SQ_G8 As Long = 97, SQ_H8 As Long = 98
'--- Move directions
Public Const SQ_UP                   As Long = 10
Public Const SQ_DOWN                 As Long = -10
Public Const SQ_RIGHT                As Long = 1
Public Const SQ_LEFT                 As Long = -1
Public Const SQ_UP_RIGHT             As Long = 11
Public Const SQ_UP_LEFT              As Long = 9
Public Const SQ_DOWN_RIGHT           As Long = -9
Public Const SQ_DOWN_LEFT            As Long = -11
'--- Files A-H
Public Const FILE_A                  As Long = 1, FILE_B As Long = 2, FILE_C As Long = 3, FILE_D As Long = 4, FILE_E As Long = 5, FILE_F As Long = 6, FILE_G As Long = 7, FILE_H As Long = 8
'--- Ranks 1-8
Public Const RANK_1                  As Long = 1, RANK_2 As Long = 2, RANK_3 As Long = 3, RANK_4 As Long = 4, RANK_5 As Long = 5, RANK_6 As Long = 6, RANK_7 As Long = 7, RANK_8 As Long = 8

Public Const MAX_MOVES               As Long = 250 ' max moves for a position
Public Const MAX_GAME_MOVES          As Long = 999
Public Const MAX_DEPTH               As Long = 150

Public Const ENPASSANT_WMOVE         As Long = 1   ' white pawn moves 2 rows > creates WEP_PIECE
Public Const ENPASSANT_BMOVE         As Long = 2   ' black pawn moves 2 rows > creates BEP_PIECE
Public Const ENPASSANT_CAPTURE       As Long = 3   ' en passant captures dummy piece WEP_PIECE or BEP_PIECE


Public Enum enumColor
  COL_WHITE = 1
  COL_BLACK = 0
  COL_NOPIECE = -1
End Enum

Public Enum enumPieceType
  NO_PIECE_TYPE = 0
  PT_PAWN = 1
  PT_KNIGHT = 2
  PT_BISHOP = 3
  PT_ROOK = 4
  PT_QUEEN = 5
  PT_KING = 6
  PT_ALL_PIECES = 7
  PIECE_TYPE_NB = 8
End Enum

Public Type TMOVE
  From             As Integer
  Target           As Integer
  Piece            As Integer
  Captured         As Integer
  EnPassant        As Integer
  Promoted         As Integer
  Castle           As Integer ' enumCastleFlag
  CapturedNumber   As Integer
  OrderValue       As Long
  SeeValue         As Long
  IsLegal          As Boolean
  IsChecking       As Boolean
End Type

Public Enum enumCastleFlag
  NO_CASTLE = 0
  WHITEOO = 1
  WHITEOOO = 2
  BLACKOO = 3
  BLACKOOO = 4
End Enum

Public Enum enumEndOfGame ' Game result
  NO_MATE = 0
  WHITE_WON = 1
  BLACK_WON = 2
  DRAW_RESULT = 3
  DRAW3REP_RESULT = 4
End Enum

Public Type TScore ' final score = mg+eg scaled by game phase
  MG As Long ' Midgame score
  EG As Long ' Endgame score
End Type

'--- Score values
Public Const MATE0                   As Long = 100000
Public Const MATE_IN_MAX_PLY         As Long = 100000 - 1000
Public Const VALUE_INFINITE          As Long = 111111
Public Const VALUE_NONE              As Long = 333333
Public Const VALUE_KNOWN_WIN         As Long = 10000

'-----------------
Public Board(MAX_BOARD)                           As Long ' Game board for all moves
Public NumPieces                                  As Long  '--- Current number of pieces at ply 0 in Pieces list
Public ColorSq(MAX_BOARD)                         As Long  '--- Squares color: COL_WHITE or COL_BLACK
Public Pieces(32)                                 As Long  '--- List of pieces: pointer to board position (Captured pieces ares set to zero during search
Public PieceColor(16)                             As Long  ' White COL_WHITE / Black COL_BLACK
Public PieceType(16)                              As Long  ' sample: maps black pawn and white pawn pieces to PT_PAWN
Public PieceCnt(16)                               As Long  ' number of pieces per piece type and color
Public Moved(MAX_BOARD)                            As Long  ' was piece moved? for castling and eval

Public File(MAX_BOARD)                            As Long
Public Rank(MAX_BOARD)                            As Long  ' Rank from black view
Public RankB(MAX_BOARD)                           As Long  ' Rank from black view  1 => 8
Public RelativeSq(COL_WHITE, MAX_BOARD)           As Long  ' sq from black view  1 => 8

Public DirectionOffset(7)                            As Long
Public KnightOffsets(7)                           As Long
Public BishopOffsets(3)                           As Long
Public RookOffsets(3)                             As Long

Public bWhiteToMove                               As Boolean  '--- side to move , false if black to move, often used
Public bCompIsWhite                               As Boolean
Public CastleFlag                                 As enumCastleFlag
Public WhiteCastled                               As enumCastleFlag
Public BlackCastled                               As enumCastleFlag
Public WPromotions(5)                             As Long  '--- list of promotion pieces
Public BPromotions(5)                             As Long
Public WKingLoc                                   As Long  ' white king location
Public BKingLoc                                   As Long  ' black king location


Public Moves(MAX_DEPTH, MAX_MOVES)       As TMOVE ' Generated moves [ply,Move]
Public EpPosArr(0 To MAX_DEPTH)                   As Long

Public Ply As Long
Public bTimeExit As Boolean
Public bEndGame As Boolean
Public GameMovesCnt As Long
Public arGameMoves(MAX_GAME_MOVES) As TMOVE
Public StartupBoard(MAX_BOARD)                    As Long ' Start Position used for copy to current board

Public MaxDistance(0 To SQ_H8, 0 To SQ_H8)        As Long
Public EmptyMove As TMOVE

'------------------------------------------------------------------------------



'---------------------------------------------------------------------------
' GenerateMoves()
' ===============
' Generates all Pseudo-legal move for a position. Check for legal moves later with CheckLegal
' if bCapturesOnly then only captures and promotions are generated.
'   if MovePickerDat(Ply).GenerateQSChecksCnt then checks are generated too. For QSearch first ply only.
'---------------------------------------------------------------------------
Public Function GenerateMoves(ByVal Ply As Long, Col As enumColor, _
                              NumMoves As Long) As Long
  Dim From As Long
  '--- Init special board with king checking positions for fast detection of checking moves
  
  NumMoves = 0
  If Col = COL_WHITE Then
    For From = SQ_A1 To SQ_H8
      Select Case Board(From)
        Case NO_PIECE, FRAME
        Case WPAWN
          ' note: FRAME has Bit 1 not set - like COL_BLACK: PieceColor() cannot be used here, returns NO_COL for EP piece
          If ((Board(From + 11) And COL_WHITE) = COL_BLACK) Then If Board(From + 11) <> FRAME Then TryMoveWPawn Ply, NumMoves, From, From + 11 ' capture right side
          If ((Board(From + 9) And COL_WHITE) = COL_BLACK) Then If Board(From + 9) <> FRAME Then TryMoveWPawn Ply, NumMoves, From, From + 9 ' capture left side
          If Board(From + 10) = NO_PIECE Then ' one row up
            If Rank(From) = 2 Then If Board(From + 20) = NO_PIECE Then TryMoveWPawn Ply, NumMoves, From, From + 20 ' two rows up
            TryMoveWPawn Ply, NumMoves, From, From + 10 ' one row up
          End If
        Case WKNIGHT
          TryMoveListKnight Ply, NumMoves, From
        Case WBISHOP
          TryMoveSliderList Ply, NumMoves, From, PT_BISHOP
        Case WROOK
          TryMoveSliderList Ply, NumMoves, From, PT_ROOK
        Case WQUEEN
          TryMoveSliderList Ply, NumMoves, From, PT_QUEEN
        Case WKING
          TryMoveListKing Ply, NumMoves, From
          ' Check castling
          If From = WKING_START Then
            If Moved(WKING_START) = 0 Then
              'o-o
              If Moved(SQ_H1) = 0 And Board(SQ_H1) = WROOK Then
                If Board(SQ_F1) = NO_PIECE And Board(SQ_G1) = NO_PIECE Then
                  CastleFlag = WHITEOO
                  TryCastleMove Ply, NumMoves, From, From + 2
                  CastleFlag = NO_CASTLE
                End If
              End If
              'o-o-o
              If Moved(SQ_A1) = 0 And Board(SQ_A1) = WROOK Then
                If Board(SQ_D1) = NO_PIECE And Board(SQ_C1) = NO_PIECE And Board(SQ_B1) = NO_PIECE Then
                  CastleFlag = WHITEOOO
                  TryCastleMove Ply, NumMoves, From, From - 2
                  CastleFlag = NO_CASTLE
                End If
              End If
            End If
          End If
      End Select

    Next

  Else

    For From = SQ_A1 To SQ_H8
      Select Case Board(From)
        Case NO_PIECE, FRAME
        Case BPAWN
          ' note: NO_PIECE has Bit 1 set like COL_WHITE
          If ((Board(From - 11) And COL_WHITE) = COL_WHITE) And Board(From - 11) <> NO_PIECE Then TryMoveBPawn Ply, NumMoves, From, From - 11
          If ((Board(From - 9) And COL_WHITE) = COL_WHITE) And Board(From - 9) <> NO_PIECE Then TryMoveBPawn Ply, NumMoves, From, From - 9
          If Board(From - 10) = NO_PIECE Then
            If Rank(From) = 7 Then If Board(From - 20) = NO_PIECE Then TryMoveBPawn Ply, NumMoves, From, From - 20
            TryMoveBPawn Ply, NumMoves, From, From - 10
          End If
        Case BKNIGHT
          TryMoveListKnight Ply, NumMoves, From
        Case BBISHOP
          TryMoveSliderList Ply, NumMoves, From, PT_BISHOP
        Case BROOK
          TryMoveSliderList Ply, NumMoves, From, PT_ROOK
        Case BQUEEN
          TryMoveSliderList Ply, NumMoves, From, PT_QUEEN
        Case BKING
          TryMoveListKing Ply, NumMoves, From
          ' Check castling
          If From = BKING_START Then
            If Moved(BKING_START) = 0 Then
              'o-o
              If Moved(SQ_H8) = 0 And Board(SQ_H8) = BROOK Then
                If Board(SQ_F8) = NO_PIECE And Board(SQ_G8) = NO_PIECE Then
                  CastleFlag = BLACKOO
                  TryCastleMove Ply, NumMoves, From, From + 2
                  CastleFlag = NO_CASTLE
                End If
              End If
              'o-o-o
              If Moved(SQ_A8) = 0 And Board(SQ_A8) = BROOK Then
                If Board(SQ_D8) = NO_PIECE And Board(SQ_C8) = NO_PIECE And Board(SQ_B8) = NO_PIECE Then
                  CastleFlag = BLACKOOO
                  TryCastleMove Ply, NumMoves, From, From - 2
                  CastleFlag = NO_CASTLE
                End If
              End If
            End If
          End If
      End Select
    Next
  End If
  
  GenerateMoves = NumMoves ' return move count
End Function

Private Function TryMoveWPawn(ByVal Ply As Long, _
                              NumMoves As Long, _
                              ByVal From As Long, _
                              ByVal Target As Long) As Boolean
  If Board(Target) = FRAME Then Exit Function
  Dim PieceFrom As Long, PieceTarget As Long
  PieceFrom = Board(From): PieceTarget = Board(Target)
  Debug.Assert PieceTarget <> FRAME

  If Rank(From) = 7 Then
      ' White Promotion
      Dim PromotePiece As Long
      For PromotePiece = 1 To 4 ' for each promotion piece type
        With Moves(Ply, NumMoves)
         .From = From: .Target = Target: .Captured = PieceTarget: .EnPassant = 0: .Castle = NO_CASTLE: .Promoted = WPromotions(PromotePiece): .Piece = .Promoted: .IsChecking = False: .IsLegal = False: .SeeValue = VALUE_NONE: .OrderValue = 0
        End With
        NumMoves = NumMoves + 1
      Next
  Else
    With Moves(Ply, NumMoves)
      Select Case PieceTarget
      Case BEP_PIECE
        .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        .EnPassant = ENPASSANT_CAPTURE: NumMoves = NumMoves + 1
      Case NO_PIECE, WEP_PIECE ' WEP_PIECE should not appear
        '---Normal move
        .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        If Target - From = 20 Then .EnPassant = ENPASSANT_WMOVE
        .IsChecking = False: NumMoves = NumMoves + 1
      Case FRAME
      Case Else
        ' Normal capture.
        .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        NumMoves = NumMoves + 1
      End Select
    End With
  End If

End Function

Private Function TryMoveBPawn(ByVal Ply As Long, _
                              NumMoves As Long, _
                              ByVal From As Long, _
                              ByVal Target As Long) As Boolean
  If Board(Target) = FRAME Then Exit Function
  Dim PieceFrom As Long, PieceTarget As Long
  PieceFrom = Board(From): PieceTarget = Board(Target)
  Debug.Assert PieceTarget <> FRAME

  If Rank(From) = 2 Then
      ' Black Promotion
      Dim PromotePiece As Long
      For PromotePiece = 1 To 4
        With Moves(Ply, NumMoves)
         .From = From: .Target = Target: .Captured = PieceTarget: .EnPassant = 0: .Castle = NO_CASTLE: .Promoted = BPromotions(PromotePiece): .Piece = .Promoted: .IsChecking = False: .IsLegal = False: .SeeValue = VALUE_NONE: .OrderValue = 0
        End With
        NumMoves = NumMoves + 1
      Next
  Else
    With Moves(Ply, NumMoves)
      Select Case PieceTarget
      Case WEP_PIECE
        .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        .EnPassant = ENPASSANT_CAPTURE: NumMoves = NumMoves + 1
      Case NO_PIECE, BEP_PIECE ' BEP_PIECE should not appear
        '--- Normal move, not a capture, promotion ---
         '---Normal move, not generated in QSearch (exception: when in check)
        .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False:  .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
         If Target - From = -20 Then .EnPassant = ENPASSANT_BMOVE
        .IsChecking = False: NumMoves = NumMoves + 1
      Case FRAME
      Case Else
        ' Normal capture.
        .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        NumMoves = NumMoves + 1
      End Select
    End With
  End If
End Function

Private Function TryMoveListKnight(ByVal Ply As Long, _
                                   NumMoves As Long, _
                                   ByVal From As Long) As Boolean
  '--- Knights only moves
  Dim Target As Long, ActDir As Long, PieceFrom As Long, PieceTarget As Long, PieceCol As Long
  PieceFrom = Board(From): PieceCol = (PieceFrom And COL_WHITE)

  For ActDir = 0 To 7
    Target = From + KnightOffsets(ActDir): PieceTarget = Board(Target)
    Select Case PieceTarget
    Case NO_PIECE, WEP_PIECE, BEP_PIECE
      '---Normal move
      With Moves(Ply, NumMoves)
        .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
      End With
      NumMoves = NumMoves + 1
    Case FRAME ' go on with next direction
    Case Else
      ' Captures
      If PieceCol <> (PieceTarget And COL_WHITE) Then ' Capture of own piece not allowed
        With Moves(Ply, NumMoves)
          .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        End With
        NumMoves = NumMoves + 1
      End If
    End Select
  Next ActDir

End Function

Private Function TryMoveListKing(ByVal Ply As Long, _
                                 NumMoves As Long, _
                                 ByVal From As Long) As Boolean
  '--- Kings only
  Dim Target As Long, ActDir As Long, PieceFrom As Long, PieceTarget As Long, PieceCol As Long

  PieceFrom = Board(From): PieceCol = (PieceFrom And COL_WHITE)

  For ActDir = 0 To 7
    Target = From + DirectionOffset(ActDir): PieceTarget = Board(Target)
    Select Case PieceTarget
    Case NO_PIECE, WEP_PIECE, BEP_PIECE
      '---Normal move
      With Moves(Ply, NumMoves)
        .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
      End With
      NumMoves = NumMoves + 1
    Case FRAME ' go on with next direction
    Case Else
      ' Captures
      If PieceCol <> (PieceTarget And COL_WHITE) Then ' Capture of own piece not allowed
        With Moves(Ply, NumMoves)
          .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        End With
        NumMoves = NumMoves + 1
      End If
    End Select
  Next ActDir

End Function

Private Function TryCastleMove(ByVal Ply As Long, _
                               NumMoves As Long, _
                               ByVal From As Long, _
                               ByVal Target As Long) As Boolean
  If Board(Target) = FRAME Then Exit Function
  Dim CurrentMove As TMOVE, PieceFrom As Long, PieceTarget As Long
  PieceFrom = Board(From): PieceTarget = Board(Target): TryCastleMove = False
  If CastleFlag <> NO_CASTLE Then
    'If Not bGenCapturesOnly Then
      CurrentMove.From = From
      CurrentMove.Target = Target
      CurrentMove.Piece = PieceFrom
      CurrentMove.Captured = PieceTarget
      CurrentMove.EnPassant = 0
      CurrentMove.Castle = CastleFlag
      CurrentMove.Promoted = 0: CurrentMove.IsChecking = False
      CurrentMove.SeeValue = VALUE_NONE
      CastleFlag = NO_CASTLE
      Moves(Ply, NumMoves) = CurrentMove
      NumMoves = NumMoves + 1
      TryCastleMove = True
    'End If
  End If
End Function

Private Sub TryMoveSliderList(ByVal Ply As Long, _
                              NumMoves As Long, _
                              ByVal From As Long, _
                              ByVal PieceType As Long)
  Dim Target As Long, ActDir As Long, Offset As Long
  Dim PieceFrom As Long, PieceTarget As Long, DirStart As Long, DirEnd As Long, PieceCol As Long

  PieceFrom = Board(From): PieceCol = (PieceFrom And COL_WHITE)

  Select Case PieceType ' get move directions
    Case PT_ROOK: DirStart = 0: DirEnd = 3 ' Rook
    Case PT_BISHOP: DirStart = 4: DirEnd = 7 ' Bishop
    Case Else: DirStart = 0: DirEnd = 7 ' Queen
  End Select

  For ActDir = DirStart To DirEnd ' for all possible directions
      Offset = DirectionOffset(ActDir): Target = From + Offset
      Do While Board(Target) <> FRAME '--- Slide loop
        PieceTarget = Board(Target)
        If PieceTarget < NO_PIECE Then ' Captures or own piece, not EnPassant pieces because NO_PIECE
          If PieceCol <> (PieceTarget And COL_WHITE) Then ' Capture of own piece not allowed, color in last bit of piece (even/uneven)
            ' Capture of opponent color: add move to list
            With Moves(Ply, NumMoves)
              .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
            End With
            NumMoves = NumMoves + 1
          End If
          Exit Do '<<< own or opp piece reached ->end for this direction
        End If ' PieceTarget < NO_PIECE

        With Moves(Ply, NumMoves) ' add move to list
          .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = NO_PIECE: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        End With
        NumMoves = NumMoves + 1

        Target = Target + Offset
      Loop  '<<< End slider loop

    Next ActDir ' next direction
End Sub


Public Function MoveText(CompMove As TMOVE) As String
  ' Returns move string for data type TMove
  ' Sample: ComPMove.from= 22: CompMove.target=24: MsgBox CompMove  >  "a2a4"
  Dim sCoordMove As String
  If CompMove.From = 0 Then MoveText = "     ": Exit Function
  sCoordMove = Chr$(File(CompMove.From) + 96) & Rank(CompMove.From)
  If CompMove.Captured <> NO_PIECE And CompMove.Captured > 0 Then sCoordMove = sCoordMove & "x"
  sCoordMove = sCoordMove & Chr$(File(CompMove.Target) + 96) & Rank(CompMove.Target)
  If CompMove.IsChecking Then sCoordMove = sCoordMove & "+"
  If CompMove.Promoted <> 0 Then

    Select Case CompMove.Promoted
      Case WKNIGHT, BKNIGHT
        sCoordMove = sCoordMove & "n"
      Case WROOK, BROOK
        sCoordMove = sCoordMove & "r"
      Case WBISHOP, BBISHOP
        sCoordMove = sCoordMove & "b"
      Case WQUEEN, BQUEEN
        sCoordMove = sCoordMove & "q"
    End Select

  End If
  MoveText = sCoordMove
End Function

Public Function Translate(ByVal isTextEN As String) As String
'  Dim i As Long
'  If pbIsOfficeMode And psLanguage = "DE" Then
'    For i = 1 To LangCnt
'      If LanguageENArr(i) = isTextEN Then Translate = LanguageArr(i): Exit Function
'    Next
'  End If
  Translate = isTextEN
End Function

Public Function FieldNumToCoord(ByVal ilFieldNum As Long) As String
  FieldNumToCoord = Chr$(Asc("a") + ((ilFieldNum - 1) Mod 8)) & Chr$(Asc("1") + ((ilFieldNum - 1) \ 8))
End Function

Public Sub ParseCommand(isCommand As String)
  If CheckLegalRootMove(isCommand) Then
    
  End If
End Sub

'---------------------------------------------------------------------------
'InCheck() Color to move in check?
'---------------------------------------------------------------------------
Public Function InCheck() As Boolean
  If bWhiteToMove Then
    InCheck = IsAttackedByB(WKingLoc)
  Else
    InCheck = IsAttackedByW(BKingLoc)
  End If
End Function

'---------------------------------------------------------------------------
'- IsAttackedByW() - square attacked by white ?  Also used for checking legal move
'---------------------------------------------------------------------------
Public Function IsAttackedByW(ByVal Location As Long) As Boolean
  Dim i        As Long, Target As Long, Offset As Long, Piece As Long
  Dim OppQRCnt As Long, OppQBCnt As Long
  IsAttackedByW = True
  OppQRCnt = PieceCnt(WQUEEN) + PieceCnt(WROOK): OppQBCnt = PieceCnt(WQUEEN) + PieceCnt(WBISHOP)

  ' vertical+horizontal: Queen, Rook, King
  For i = 0 To 3
    Offset = DirectionOffset(i): Target = Location + Offset: Piece = Board(Target)
    If Piece <> FRAME Then
      If Piece = WKING Then Exit Function
      If OppQRCnt > 0 Then

        Do While Piece <> FRAME
          If Piece < NO_PIECE Then If Piece = WROOK Or Piece = WQUEEN Then Exit Function Else Exit Do
          Target = Target + Offset: Piece = Board(Target)
        Loop

      End If
    End If
  Next

  ' diagonal: Queen, Bishop, Pawn, King
  For i = 4 To 7
    Offset = DirectionOffset(i): Target = Location + Offset: Piece = Board(Target)
    If Piece <> FRAME Then
      If Piece = WPAWN Then
        If ((i = 5) Or (i = 7)) Then Exit Function
      ElseIf Piece = WKING Then Exit Function
      ElseIf OppQBCnt <> 0 Then

        Do While Piece <> FRAME
          If Piece < NO_PIECE Then If Piece = WBISHOP Or Piece = WQUEEN Then Exit Function Else Exit Do
          Target = Target + Offset: Piece = Board(Target)
        Loop

      End If
    End If
  Next

  If PieceCnt(WKNIGHT) > 0 Then
    For i = 0 To 7
      If Board(Location + KnightOffsets(i)) = WKNIGHT Then Exit Function ' Knight
    Next
  End If
  IsAttackedByW = False
End Function


'---------------------------------------------------------------------------
'- IsAttackedByB() - square attacked by black ?  Also used for checking legal move
'---------------------------------------------------------------------------
Public Function IsAttackedByB(ByVal Location As Long) As Boolean
  Dim i        As Long, Target As Long, Offset As Long, Piece As Long
  Dim OppQRCnt As Long, OppQBCnt As Long
  IsAttackedByB = True
  OppQRCnt = PieceCnt(BQUEEN) + PieceCnt(BROOK): OppQBCnt = PieceCnt(BQUEEN) + PieceCnt(BBISHOP)

  ' vertical+horizontal: Queen, Rook, King
  For i = 0 To 3
    Offset = DirectionOffset(i): Target = Location + Offset: Piece = Board(Target)
    If Piece <> FRAME Then
      If Piece = BKING Then Exit Function
      If OppQRCnt > 0 Then

        Do While Piece <> FRAME
          If Piece < NO_PIECE Then If Piece = BROOK Or Piece = BQUEEN Then Exit Function Else Exit Do
          Target = Target + Offset: Piece = Board(Target)
        Loop

      End If
    End If
  Next

  ' diagonal: Queen, Bishop, Pawn, King
  For i = 4 To 7
    Offset = DirectionOffset(i): Target = Location + Offset: Piece = Board(Target)
    If Piece <> FRAME Then
      If Piece = BPAWN Then
        If ((i = 4) Or (i = 6)) Then Exit Function
      ElseIf Piece = BKING Then Exit Function
      ElseIf OppQBCnt <> 0 Then

        Do While Piece <> FRAME
          If Piece < NO_PIECE Then If Piece = BBISHOP Or Piece = BQUEEN Then Exit Function Else Exit Do
          Target = Target + Offset: Piece = Board(Target)
        Loop

      End If
    End If
  Next

  If PieceCnt(BKNIGHT) > 0 Then
    For i = 0 To 7
      If Board(Location + KnightOffsets(i)) = BKNIGHT Then Exit Function ' Knight
    Next
  End If
  IsAttackedByB = False
End Function

'---------------------------------------------------------------------------
' Board File character to number  A => 1
'---------------------------------------------------------------------------
Public Function FileRev(ByVal sFile As String) As Long
  FileRev = Asc(LCase$(sFile)) - 96
End Function

'---------------------------------------------------------------------------
'RankRev() - Board Rank number to square number Rank 2 = 30
'---------------------------------------------------------------------------
Public Function RankRev(ByVal sRank As String) As Long
  RankRev = (Val(sRank) + 1) * 10
End Function

Public Sub RemoveEpPiece()
  ' Remove EP from Previous Move
  If EpPosArr(Ply) > 0 Then Board(EpPosArr(Ply)) = NO_PIECE
End Sub

Public Sub ResetEpPiece()
  ' Reset EP from Previous Move
  If EpPosArr(Ply) > 0 Then
    Select Case Rank(EpPosArr(Ply))
      Case 3
        Board(EpPosArr(Ply)) = WEP_PIECE
      Case 6
        Board(EpPosArr(Ply)) = BEP_PIECE
    End Select
  End If
End Sub

Public Sub CleanEpPieces()
  Dim i As Long

  For i = SQ_A1 To SQ_H8
    If Board(i) = WEP_PIECE Or Board(i) = BEP_PIECE Then Board(i) = NO_PIECE
  Next

End Sub

Public Sub InitBoardColors()
  Dim x As Long, y As Long, ColSq  As Long, IsWhite As Boolean

  For y = 1 To 8
    IsWhite = CBool((y And 1) = 0)

    For x = 1 To 8
      If IsWhite Then ColSq = COL_WHITE Else ColSq = COL_BLACK
      ColorSq(20 + x + (y - 1) * 10) = ColSq
      IsWhite = Not IsWhite
    Next
  Next

End Sub


Public Function GenerateLegalMoves(olTotalMoves As Long) As Long
  ' Returns all moves in Moves(ply). Moves(x).IsLegal=true for legal moves
  Dim LegalMoves As Long, lLegalMoves As Long, i As Long, NumMoves As Long
  Dim Col As enumColor
  If bWhiteToMove Then Col = COL_WHITE Else Col = COL_BLACK
  
  GenerateMoves Ply, Col, NumMoves
  Ply = 1: lLegalMoves = 0
  
  For i = 0 To NumMoves - 1
    RemoveEpPiece
    MakeMove Moves(Ply, i)
    If CheckLegal(Moves(Ply, i)) Then
     Moves(Ply, i).IsLegal = True: lLegalMoves = lLegalMoves + 1
     'Debug.Print MoveText(Moves(Ply, i))
    End If
    UnmakeMove Moves(Ply, i)
    ResetEpPiece
    'Debug.Print MovesText(Moves(ply, i)), Moves(Ply, i).IsLegal
  Next
  olTotalMoves = NumMoves
  GenerateLegalMoves = lLegalMoves
End Function


Public Sub ShoCOL_WHITEToMove()
  With frmChessX.lblColToMove
    If bWhiteToMove Then
      .BackColor = vbWhite
      .ForeColor = vbBlack
      .Caption = Translate("White to move")
    Else
      .BackColor = vbBlack
      .ForeColor = vbWhite
      .Caption = Translate("Black to move")
    End If
  End With
End Sub

Public Sub ShowLastMoveAtBoard(Move As TMOVE)
 ShowMove Move.From, Move.Target
End Sub

Public Sub ShowMove(From As Integer, Target As Integer)
 ' show move on board with different backcolor
 Dim Pos As Long, ctrl As Control
 
 If From > 0 Then
    For Each ctrl In frmChessX.Controls
      Pos = Val("0" & ctrl.Tag)
      If Pos = From Then ctrl.BackColor = &HC0FFC0
    Next ctrl
 End If
 
 If Target > 0 Then
    For Each ctrl In frmChessX.Controls
      Pos = Val("0" & ctrl.Tag)
      If Pos = Target Then ctrl.BackColor = &HC0FFC0
    Next ctrl
 End If
End Sub

'---------------------------------------------------------------------------
' MS Excel: init german translation from internal worksheet
'---------------------------------------------------------------------------
Public Function InitTranslateExcel() As Boolean

End Function



'---------------------------------------------------------------------------
' InitEngine()
'---------------------------------------------------------------------------
Public Sub InitEngine()
  '------------------------------
  '--- init arrays
  '------------------------------
'  Erase PVLength()
'  Erase PV()
'  Erase History()
'  Erase CaptureHistory()
'  Erase CounterMove()
'  Erase ContinuationHistory()
  
'  Erase Pieces()
'  Erase Squares()
'  Erase Killer()
  Erase Board()
'  Erase Moved()
'  Erase MovesList()
  Erase arGameMoves()
'  Erase GamePosHash()
  
  
  InitRankFile
  InitPieceColor
  InitPieceTypes
  InitBoard ' fill board with NO_PIECE, FRAME = 0
  
  
  '-------------------------------------
  '--- move offsets  ---
  '-------------------------------------
  ' direction index 0-3: Orthogonal (Queen+Rook), 4-7=diagonal (Queen+Bishop)
  ReadIntArr DirectionOffset(), SQ_UP, SQ_DOWN, SQ_RIGHT, SQ_LEFT, SQ_UP_RIGHT, SQ_DOWN_LEFT, SQ_UP_LEFT, SQ_DOWN_RIGHT
  ReadIntArr KnightOffsets(), 8, 19, 21, 12, -8, -19, -21, -12
  ReadIntArr BishopOffsets(), SQ_UP_LEFT, SQ_UP_RIGHT, SQ_DOWN_RIGHT, SQ_DOWN_LEFT
  ReadIntArr RookOffsets(), SQ_RIGHT, SQ_LEFT, SQ_UP, SQ_DOWN
 ' OppositeDir(SQ_RIGHT) = SQ_LEFT: OppositeDir(SQ_LEFT) = SQ_RIGHT: OppositeDir(SQ_UP) = SQ_DOWN: OppositeDir(SQ_DOWN) = SQ_UP
 ' OppositeDir(SQ_UP_RIGHT) = SQ_DOWN_LEFT: OppositeDir(SQ_DOWN_LEFT) = SQ_UP_RIGHT: OppositeDir(SQ_UP_LEFT) = SQ_DOWN_RIGHT: OppositeDir(SQ_DOWN_RIGHT) = SQ_UP_LEFT
  
  ReadIntArr WPromotions(), 0, WQUEEN, WROOK, WKNIGHT, WBISHOP
  ReadIntArr BPromotions(), 0, BQUEEN, BROOK, BKNIGHT, BBISHOP
  ReadIntArr PieceType, 0, PT_PAWN, PT_PAWN, PT_KNIGHT, PT_KNIGHT, PT_BISHOP, PT_BISHOP, PT_ROOK, PT_ROOK, PT_QUEEN, PT_QUEEN, PT_KING, PT_KING, NO_PIECE_TYPE, PT_PAWN, PT_PAWN
  InitRankFile ' must be before InitMaxDistance
  InitBoardColors
  InitMaxDistance
 ' InitSqBetween
 ' InitSameXRay
 ' InitAttackBitCnt

  ' setup empty move
  With EmptyMove
    .From = 0: .Target = 0: .Piece = NO_PIECE: .Castle = NO_CASTLE: .Promoted = 0: .Captured = NO_PIECE: .CapturedNumber = 0
    .EnPassant = 0: .IsChecking = False: .IsLegal = False: .OrderValue = 0: .SeeValue = VALUE_NONE
  End With

  '--------------------------------------------
  '--- startup board
  '--------------------------------------------
  ReadIntArr StartupBoard(), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, WROOK, WKNIGHT, WBISHOP, WQUEEN, WKING, WBISHOP, WKNIGHT, WROOK, 0, 0, WPAWN, WPAWN, WPAWN, WPAWN, WPAWN, WPAWN, WPAWN, WPAWN, 0, 0, 13, 13, 13, 13, 13, 13, 13, 13, 0, 0, 13, 13, 13, 13, 13, 13, 13, 13, 0, 0, 13, 13, 13, 13, 13, 13, 13, 13, 0, 0, 13, 13, 13, 13, 13, 13, 13, 13, 0, 0, BPAWN, BPAWN, BPAWN, BPAWN, BPAWN, BPAWN, BPAWN, BPAWN, 0, 0, BROOK, BKNIGHT, BBISHOP, BQUEEN, BKING, BBISHOP, BKNIGHT, BROOK, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
  
 
  ' Init EPD table
'  InitEPDTable
 ' bUseBook = InitBook
  ' Init Hash
 ' InitZobrist
  ' Endgame tablebase access (via online web service or fathom.exe)
 ' InitTableBases
  ' Init game
'  InitGame
End Sub

Public Sub SendCommand(sCommand As String)

End Sub

Public Sub MakeMove(mMove As TMOVE)
  '--- Do move on board
  Dim From      As Long, Target As Long
  Dim Captured  As Long, EnPassant As Long
  Dim Promoted  As Long, Castle As Long
  Dim PieceFrom As Long

  With mMove
    From = .From: Target = .Target: Captured = .Captured: EnPassant = .EnPassant: Promoted = .Promoted: Castle = .Castle
  End With

  PieceFrom = Board(From)
  Board(From) = NO_PIECE: Moved(From) = Moved(From) + 1
  'mMove.CapturedNumber = Squares(Target)
  'Pieces(Squares(From)) = Target: Pieces(Squares(Target)) = 0
  'Squares(Target) = Squares(From): Squares(From) = 0
  'If PieceFrom = WPAWN Or PieceFrom = BPAWN Or Board(Target) < NO_PIECE Or Promoted <> 0 Then Fifty = 0 Else Fifty = Fifty + 1
  
  ' En Passant
  EpPosArr(Ply + 1) = 0
  If EnPassant <> 0 Then
    If EnPassant = ENPASSANT_WMOVE Then
      Board(From + 10) = WEP_PIECE
      EpPosArr(Ply + 1) = From + 10
    ElseIf EnPassant = ENPASSANT_BMOVE Then
      Board(From - 10) = BEP_PIECE
      EpPosArr(Ply + 1) = From - 10
    End If
    If EnPassant = ENPASSANT_CAPTURE Then '--- EP capture move
      If PieceFrom = WPAWN Then
        Board(Target) = PieceFrom
        Board(Target - 10) = NO_PIECE: PieceCntMinus BPAWN
       ' mMove.CapturedNumber = Squares(Target - 10)
       ' Pieces(Squares(Target - 10)) = 0: Squares(Target - 10) = 0
      ElseIf PieceFrom = BPAWN Then
        Board(Target) = PieceFrom
        Board(Target + 10) = NO_PIECE: PieceCntMinus WPAWN
       ' mMove.CapturedNumber = Squares(Target + 10)
       ' Pieces(Squares(Target + 10)) = 0: Squares(Target + 10) = 0
      End If
      GoTo lblExit
    End If
  End If
  'Castle: additional rook move here, King later as normal move
  If Castle <> NO_CASTLE Then

    Select Case Castle
      Case WHITEOO
        WhiteCastled = WHITEOO
        Board(SQ_H1) = NO_PIECE: Moved(SQ_H1) = Moved(SQ_H1) + 1
        Board(SQ_F1) = WROOK: Moved(SQ_F1) = Moved(SQ_F1) + 1
       ' Pieces(Squares(SQ_H1)) = SQ_F1: Squares(SQ_F1) = Squares(SQ_H1): Squares(SQ_H1) = 0
        Board(SQ_G1) = WKING: Moved(SQ_G1) = Moved(SQ_G1) + 1: WKingLoc = SQ_G1
        GoTo lblExit
      Case WHITEOOO
        WhiteCastled = WHITEOOO
        Board(SQ_A1) = NO_PIECE: Moved(SQ_A1) = Moved(SQ_A1) + 1
        Board(SQ_D1) = WROOK: Moved(SQ_D1) = Moved(SQ_D1) + 1
      '  Pieces(Squares(SQ_A1)) = SQ_D1: Squares(SQ_D1) = Squares(SQ_A1): Squares(SQ_A1) = 0
        Board(SQ_C1) = WKING: Moved(SQ_C1) = Moved(SQ_C1) + 1: WKingLoc = SQ_C1
        GoTo lblExit
      Case BLACKOO
        BlackCastled = BLACKOO
        Board(SQ_H8) = NO_PIECE: Moved(SQ_H8) = Moved(SQ_H8) + 1
        Board(SQ_F8) = BROOK: Moved(SQ_F8) = Moved(SQ_F8) + 1
      '  Pieces(Squares(SQ_H8)) = SQ_F8: Squares(SQ_F8) = Squares(SQ_H8): Squares(SQ_H8) = 0
        Board(SQ_G8) = BKING: Moved(SQ_G8) = Moved(SQ_G8) + 1: BKingLoc = SQ_G8
        GoTo lblExit
      Case BLACKOOO
        BlackCastled = BLACKOOO
        Board(SQ_A8) = NO_PIECE: Moved(SQ_A8) = Moved(SQ_A8) + 1
        Board(SQ_D8) = BROOK: Moved(SQ_D8) = Moved(SQ_D8) + 1
       ' Pieces(Squares(SQ_A8)) = SQ_D8: Squares(SQ_D8) = Squares(SQ_A8): Squares(SQ_A8) = 0
        Board(SQ_C8) = BKING: Moved(SQ_C8) = Moved(SQ_C8) + 1: BKingLoc = SQ_C8
        GoTo lblExit
    End Select

  End If
  If Promoted <> 0 Then
    PieceCntPlus Promoted
    Board(Target) = Promoted
    PieceCntMinus PieceFrom
    Moved(Target) = Moved(Target) + 1
  Else

    '--- normal move
    Select Case PieceFrom
      Case WKING: WKingLoc = Target
      Case BKING: BKingLoc = Target
    End Select

    Board(Target) = PieceFrom: Moved(Target) = Moved(Target) + 1
  End If
  If Captured > 0 Then If Captured < NO_PIECE Then PieceCntMinus Captured
lblExit:
  bWhiteToMove = Not bWhiteToMove
End Sub

Public Sub UnmakeMove(mMove As TMOVE)
  ' take back this move on board
  Dim From     As Long, Target As Long
  Dim Captured As Long, EnPassant As Long, CapturedNumber As Long
  Dim Promoted As Long, Castle As Long, PieceTarget As Long

  With mMove
    From = .From: Target = .Target: Captured = .Captured
    EnPassant = .EnPassant: Promoted = .Promoted: Castle = .Castle: CapturedNumber = .CapturedNumber
  End With

  PieceTarget = Board(Target)
  'Squares(From) = Squares(Target): Squares(Target) = CapturedNumber
  'Pieces(Squares(Target)) = Target: Pieces(Squares(From)) = From
  'Fifty = arFiftyMove(Ply)
  If Castle <> NO_CASTLE Then

    Select Case Castle
      Case WHITEOO
        WhiteCastled = NO_CASTLE
        Board(SQ_F1) = NO_PIECE: Moved(SQ_F1) = Moved(SQ_F1) - 1
        Board(SQ_H1) = WROOK: Moved(SQ_H1) = Moved(SQ_H1) - 1
       ' Squares(SQ_H1) = Squares(SQ_F1): Squares(SQ_F1) = 0: Pieces(Squares(SQ_H1)) = SQ_H1
        Board(SQ_E1) = WKING: Moved(SQ_E1) = Moved(SQ_E1) - 1: WKingLoc = SQ_E1
        Board(SQ_G1) = NO_PIECE: Moved(SQ_G1) = Moved(SQ_G1) - 1
        GoTo lblExit
      Case WHITEOOO
        WhiteCastled = NO_CASTLE
        Board(SQ_D1) = NO_PIECE: Moved(SQ_D1) = Moved(SQ_D1) - 1
        Board(SQ_A1) = WROOK: Moved(SQ_A1) = Moved(SQ_A1) - 1
       ' Squares(SQ_A1) = Squares(SQ_D1): Squares(SQ_D1) = 0: Pieces(Squares(SQ_A1)) = SQ_A1
        Board(SQ_E1) = WKING: Moved(SQ_E1) = Moved(SQ_E1) - 1: WKingLoc = SQ_E1
        Board(SQ_C1) = NO_PIECE: Moved(SQ_C1) = Moved(SQ_C1) - 1
        GoTo lblExit
      Case BLACKOO
        BlackCastled = NO_CASTLE
        Board(SQ_F8) = NO_PIECE: Moved(SQ_F8) = Moved(SQ_F8) - 1
        Board(SQ_H8) = BROOK: Moved(SQ_H8) = Moved(SQ_H8) - 1
       ' Squares(SQ_H8) = Squares(SQ_F8): Squares(SQ_F8) = 0: Pieces(Squares(SQ_H8)) = SQ_H8
        Board(SQ_E8) = BKING: Moved(SQ_E8) = Moved(SQ_E8) - 1: BKingLoc = SQ_E8
        Board(SQ_G8) = NO_PIECE: Moved(SQ_G8) = Moved(SQ_G8) - 1
        GoTo lblExit
      Case BLACKOOO
        BlackCastled = NO_CASTLE
        Board(SQ_D8) = NO_PIECE: Moved(SQ_D8) = Moved(SQ_D8) - 1
        Board(SQ_A8) = BROOK: Moved(SQ_A8) = Moved(SQ_A8) - 1
       ' Squares(SQ_A8) = Squares(SQ_D8): Squares(SQ_D8) = 0: Pieces(Squares(SQ_A8)) = SQ_A8
        Board(SQ_E8) = BKING: Moved(SQ_E8) = Moved(SQ_E8) - 1: BKingLoc = SQ_E8
        Board(SQ_C8) = NO_PIECE: Moved(SQ_C8) = Moved(SQ_C8) - 1
        GoTo lblExit
    End Select

  End If
  If EnPassant <> 0 Then
    If EnPassant = ENPASSANT_WMOVE Then
      Board(From + 10) = NO_PIECE
    ElseIf EnPassant = ENPASSANT_BMOVE Then
      Board(From - 10) = NO_PIECE
    End If
    If EnPassant = ENPASSANT_CAPTURE Then
      If PieceTarget = WPAWN Then
        Board(From) = PieceTarget
        Board(Target) = BEP_PIECE
        Board(Target - 10) = BPAWN: PieceCntPlus BPAWN
      '  Squares(Target - 10) = CapturedNumber
       ' Pieces(CapturedNumber) = Target - 10
       ' Squares(Target) = 0
      ElseIf PieceTarget = BPAWN Then
        Board(From) = PieceTarget
        Board(Target) = WEP_PIECE
        Board(Target + 10) = WPAWN: PieceCntPlus WPAWN
       ' Squares(Target + 10) = CapturedNumber
       ' Pieces(CapturedNumber) = Target + 10
       ' Squares(Target) = 0
      End If
      Moved(From) = Moved(From) - 1
      GoTo lblExit
    End If
  End If
  If Promoted <> 0 Then
    If (Promoted And 1) = COL_WHITE Then
      Board(From) = WPAWN: PieceCntPlus WPAWN
      PieceCntMinus Board(Target)
      Board(Target) = Captured
      Moved(From) = Moved(From) - 1
      Moved(Target) = Moved(Target) - 1
    Else
      Board(From) = BPAWN: PieceCntPlus BPAWN
      PieceCntMinus Board(Target)
      Board(Target) = Captured
      Moved(From) = Moved(From) - 1
      Moved(Target) = Moved(Target) - 1
    End If
  Else

    '--- normal move
    Select Case PieceTarget
      Case WKING: WKingLoc = From
      Case BKING: BKingLoc = From
    End Select

    Board(From) = PieceTarget: Moved(From) = Moved(From) - 1
    Board(Target) = Captured: Moved(Target) = Moved(Target) - 1
  End If
  If Captured > 0 Then If Captured < NO_PIECE Then PieceCntPlus Captured
lblExit:
  
  bWhiteToMove = Not bWhiteToMove ' switch side to move
  
End Sub
'---------------------------------------------------------------------------
'- CheckLegal() - Legal move?
'---------------------------------------------------------------------------
Public Function CheckLegal(mMove As TMOVE) As Boolean
  CheckLegal = False
  If mMove.From < SQ_A1 Then Exit Function
  If mMove.Castle = NO_CASTLE Then
    If bWhiteToMove Then
      If IsAttackedByW(BKingLoc) Then Exit Function ' BKing mate?
    Else
      If IsAttackedByB(WKingLoc) Then Exit Function ' WKing mate?
    End If
  Else

    ' Castling
    Select Case mMove.Castle
      Case WHITEOO:
        If IsAttackedByB(WKING_START) Then Exit Function
        If IsAttackedByB(WKING_START + 1) Then Exit Function
        If IsAttackedByB(WKING_START + 2) Then Exit Function
      Case WHITEOOO:
        If IsAttackedByB(WKING_START) Then Exit Function
        If IsAttackedByB(WKING_START - 1) Then Exit Function
        If IsAttackedByB(WKING_START - 2) Then Exit Function
      Case BLACKOO:
        If IsAttackedByW(BKING_START) Then Exit Function
        If IsAttackedByW(BKING_START + 1) Then Exit Function
        If IsAttackedByW(BKING_START + 2) Then Exit Function
      Case BLACKOOO:
        If IsAttackedByW(BKING_START) Then Exit Function
        If IsAttackedByW(BKING_START - 1) Then Exit Function
        If IsAttackedByW(BKING_START - 2) Then Exit Function
    End Select

  End If
  CheckLegal = True
End Function

Public Function CheckLegalRootMove(ByVal isMove As String) As Boolean
  Dim PlayerMove As TMOVE, i As Long, iNumMoves As Long, sCoordMove As String, sActMove As String, bLegalInput As Boolean
  'Dim Hashkey    As THashKey
  Dim sInput(4) As String, Col As enumColor
  CheckLegalRootMove = False
  If Len(Trim$(isMove)) < 4 Then Exit Function

  For i = 0 To 4
    sInput(i) = Mid$(isMove, i + 1, 1)
  Next
  If bWhiteToMove Then Col = COL_WHITE Else Col = COL_BLACK
  sActMove = Trim$(Left(isMove, Len(isMove) - 1)) ' remove vbCrLf at end
  bLegalInput = False
  '--- normal move like with 4 char: e2e4 ---
  If Not IsNumeric(sInput(0)) And IsNumeric(sInput(1)) And Not IsNumeric(sInput(2)) And IsNumeric(sInput(3)) Then
    Ply = 0
    GenerateMoves Ply, Col, iNumMoves
    PlayerMove.From = FileRev(sInput(0)) + RankRev(sInput(1))
    PlayerMove.Target = FileRev(sInput(2)) + RankRev(sInput(3))

    ' legal move?
    For i = 0 To iNumMoves - 1
      sCoordMove = CompToCoord(Moves(Ply, i))
      'Debug.Print sCoordMove
      If Trim(sActMove) = sCoordMove Then
        RemoveEpPiece
        MakeMove Moves(Ply, i)
        If CheckLegal(Moves(Ply, i)) Then
          bLegalInput = True
          PlayerMove.Captured = Moves(Ply, i).Captured
          PlayerMove.Piece = Moves(Ply, i).Piece
          PlayerMove.Promoted = Moves(Ply, i).Promoted
          PlayerMove.EnPassant = Moves(Ply, i).EnPassant
          PlayerMove.Castle = Moves(Ply, i).Castle
          PlayerMove.CapturedNumber = Moves(Ply, i).CapturedNumber
        End If
        UnmakeMove Moves(Ply, i)
        ResetEpPiece
        If bLegalInput Then Exit For
      End If
    Next

    If Not bLegalInput Then
     ' If bWinboardTrace Then LogWrite "Illegal move: " & sCoordMove
    Else
      ' do game move
       CleanEpPieces
       MakeMove PlayerMove
       PrevLastMove = LastMove
       LastMove = PlayerMove
       frmChessX.txtIO.Text = "move " & MoveText(PlayerMove)
'      PlayMove PlayerMove
'      HashBoard Hashkey, EmptyMove
'      If Is3xDraw(Hashkey, GameMovesCnt, 0) Then
'        ' Result = DRAW3REP_RESULT
'        If bWinboardTrace Then LogWrite "ParseCommand: Return Draw3Rep"
'        'SendCommand "1/2-1/2 {Draw by repetition}"
'      End If
 '     GameMovesAdd PlayerMove
      'LogWrite "move: " & sCoordMove
    End If
  End If
  CheckLegalRootMove = bLegalInput
End Function

Public Function LocCoord(ByVal Square As Long) As String
  LocCoord = UCase$(Chr$(File(Square) + 96) & Rank(Square))
End Function



Public Function RelativeRank(ByVal Col As enumColor, ByVal sq As Long) As Long
  If Col = COL_WHITE Then
    RelativeRank = Rank(sq)
  Else
    RelativeRank = (9 - Rank(sq))
  End If
End Function

'---------------------------------------------------------------------------
'CompToCoord(): Convert internal move to text output
'---------------------------------------------------------------------------
Public Function CompToCoord(CompMove As TMOVE) As String
  Dim sCoordMove As String
  If CompMove.From = 0 Then CompToCoord = "": Exit Function
  sCoordMove = Chr$(File(CompMove.From) + 96) & Rank(CompMove.From) & Chr$(File(CompMove.Target) + 96) & Rank(CompMove.Target)
  If CompMove.Promoted <> 0 Then

    Select Case CompMove.Promoted
      Case WKNIGHT, BKNIGHT
        sCoordMove = sCoordMove & "n"
      Case WROOK, BROOK
        sCoordMove = sCoordMove & "r"
      Case WBISHOP, BBISHOP
        sCoordMove = sCoordMove & "b"
      Case WQUEEN, BQUEEN
        sCoordMove = sCoordMove & "q"
    End Select

  End If
  CompToCoord = sCoordMove
End Function

Public Sub CopyIntArr(SourceArr() As Long, DestArr() As Long)
  Dim i As Long
  For i = LBound(SourceArr) To UBound(SourceArr) - 1: DestArr(i) = SourceArr(i): Next
End Sub

Public Sub InitGame()
  ' Init start position
  CopyIntArr StartupBoard, Board
  Erase Moved()
  GameMovesCnt = 0: Erase arGameMoves()
 ' HintMove = EmptyMove
  
'  InitHash
  InitPieceCnt
  'Result = NO_MATE
  bWhiteToMove = True
  bCompIsWhite = False
  WKingLoc = WKING_START
  BKingLoc = BKING_START
  WhiteCastled = NO_CASTLE
  BlackCastled = NO_CASTLE
  
  ' Init Bitboards
  Eval
End Sub

Public Sub InitMaxDistance()
  ' Max distance x or y
  Dim i As Long, j As Long
  Dim d As Long, V As Long

  For i = SQ_A1 To SQ_H8
    For j = SQ_A1 To SQ_H8
      V = Abs(Rank(i) - Rank(j))
      d = Abs(File(i) - File(j))
      If d > V Then V = d
      MaxDistance(i, j) = V
    Next j
  Next i
End Sub

Public Sub PieceCntPlus(ByVal Piece As Long)
  If Piece > FRAME And Piece < NO_PIECE Then
    PieceCnt(Piece) = PieceCnt(Piece) + 1
   ' If Piece > BPAWN And Piece < WKING Then ' King not counted
   '   If CBool(Piece And 1) Then WNonPawnPieces = WNonPawnPieces + 1 Else BNonPawnPieces = BNonPawnPieces + 1
   ' End If
  End If
End Sub

Public Sub PieceCntMinus(ByVal Piece As Long)
  If Piece > FRAME And Piece < NO_PIECE Then
    PieceCnt(Piece) = PieceCnt(Piece) - 1
   ' If Piece > BPAWN And Piece < WKING Then
   '   If CBool(Piece And 1) Then WNonPawnPieces = WNonPawnPieces - 1 Else BNonPawnPieces = BNonPawnPieces - 1
   ' End If
  End If
  Debug.Assert PieceCnt(Piece) >= 0
End Sub

'---------------------------------------------------------------------------
' InitPieceCnt
'---------------------------------------------------------------------------
Public Sub InitPieceCnt()
  Dim i As Long

  For i = SQ_A1 To SQ_H8
    If (Board(i) <> FRAME And Board(i) < NO_PIECE) Then
      PieceCntPlus Board(i)
      Select Case Board(i)
        Case WKING: WKingLoc = i
        Case BKING: BKingLoc = i
      End Select
    End If
  Next

End Sub

Public Sub InitRankFile()
  Dim Square As Long

  For Square = 1 To MAX_BOARD
    Rank(Square) = (Square \ 10) - 1
    RankB(Square) = 9 - Rank(Square)
    File(Square) = Square Mod 10
    RelativeSq(COL_WHITE, Square) = Square
    RelativeSq(COL_BLACK, Square) = SQ_A1 - 1 + File(Square) + (8 - Rank(Square)) * 10
  Next
End Sub

Public Function Piece2Alpha(ByVal iPiece As Long) As String

  Select Case iPiece
    Case WPAWN
      Piece2Alpha = "P"
    Case BPAWN
      Piece2Alpha = "p"
    Case WKNIGHT
      Piece2Alpha = "N"
    Case BKNIGHT
      Piece2Alpha = "n"
    Case WBISHOP
      Piece2Alpha = "B"
    Case BBISHOP
      Piece2Alpha = "b"
    Case WROOK
      Piece2Alpha = "R"
    Case BROOK
      Piece2Alpha = "r"
    Case WQUEEN
      Piece2Alpha = "Q"
    Case BQUEEN
      Piece2Alpha = "q"
    Case WKING
      Piece2Alpha = "K"
    Case BKING
      Piece2Alpha = "k"
    Case Else
      Piece2Alpha = "."
  End Select

End Function

Public Sub InitPieceTypes() ' assign each white and black piece a type
  ReadIntArr PieceType, 0, PT_PAWN, PT_PAWN, PT_KNIGHT, PT_KNIGHT, PT_BISHOP, PT_BISHOP, PT_ROOK, PT_ROOK, PT_QUEEN, PT_QUEEN, PT_KING, PT_KING, NO_PIECE_TYPE, PT_PAWN, PT_PAWN
End Sub

Public Function ReadIntArr(pDest() As Long, ParamArray pSrc())
  ' Read paramter list into array of type Integer
  Dim i As Long

  For i = 0 To UBound(pSrc): pDest(i) = pSrc(i): Next
End Function

Public Function ReadLngArr(pDest() As Long, ParamArray pSrc())
  ' Read paramter list into array of type Long
  Dim i As Long

  For i = 0 To UBound(pSrc): pDest(i) = pSrc(i): Next
End Function

Public Sub InitPieceColor()
  Dim Piece As Long, PieceCol As Long

  For Piece = 0 To 16
    If Piece < 1 Or Piece >= NO_PIECE Then
      PieceCol = COL_NOPIECE ' NO_PIECE, or EP-PIECE  or FRAME
    Else
      If Piece Mod 2 = COL_WHITE Then PieceCol = COL_WHITE Else PieceCol = COL_BLACK
    End If
    PieceColor(Piece) = PieceCol
  Next

End Sub


'---------------------------------------------------------------------------
'PrintPos() - board position in ASCII table
'---------------------------------------------------------------------------
Public Function PrintPos() As String
  Dim a      As Long, b As Long, c As Long
  Dim sBoard As String
  sBoard = vbCrLf
  If True Then ' Not bCompIsWhite Then  'punto di vista del B (engine e' N)
    sBoard = sBoard & " ------------------" & vbCrLf
    For a = 1 To 8
      sBoard = sBoard & (9 - a) & "| "

      For b = 1 To 8
        c = 100 - (a * 10) + b
        sBoard = sBoard & Piece2Alpha(Board(c)) & " "
      Next

      sBoard = sBoard & "| " & vbCrLf
    Next

  Else

    For a = 1 To 8
      sBoard = sBoard & a & vbTab

      For b = 1 To 8
        c = 10 + (a * 10) - b
        sBoard = sBoard & Piece2Alpha(Board(c)) & " "
      Next

      sBoard = sBoard & vbCrLf
    Next

  End If
 sBoard = sBoard & " ------------------" & vbCrLf
  sBoard = sBoard & " " & vbTab & " A B C D E F G H" & vbCrLf
  PrintPos = sBoard
End Function

Public Sub InitBoard()
  Dim Square As Long
  For Square = SQ_A1 To SQ_H8
    If File(Square) >= FILE_A And File(Square) <= FILE_H Then Board(Square) = NO_PIECE
  Next
End Sub




