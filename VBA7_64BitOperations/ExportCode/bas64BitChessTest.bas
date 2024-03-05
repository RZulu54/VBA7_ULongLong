Attribute VB_Name = "bas64BitChessTest"
'==========================================================================================
'= 64bit chess test functions
'= bitboard based evaluation function (some basics only to show the use of 64 bit LongLong)
'= (by Roger Zuehlsdorf 2024 / email:rogzuehlsdorf@yahoo.de)
'==========================================================================================
Option Explicit

#If VBA7 And Win64 Then 'Note: Win64 = Office64 bit (not Windows 64 bit)


'=============================================
'--------- 64 bit ----------------------------
'=============================================
Public BoardLL                      As LongLong ' bitboard for current board with all pieces (BKING = max index)
Public PiecesLL(COL_WHITE, PT_ALL_PIECES)  As LongLong ' bitboard for each piece type for a color on current board
Public PiecesForColLL(COL_WHITE)    As LongLong ' bitboard for all pieces of white/black on current board

Public SqToBit(MAX_BOARD) As Long  ' returns array index of a square for a 64 bit position
Public BitToSq(63) As Long         ' returns bit position for a square array index

Public PieceMovesLL(BKING, 63)      As LongLong ' bitboard for moves of each piece type at a given square on board
Public FileA_LL As LongLong, FileB_LL As LongLong, FileC_LL As LongLong, FileD_LL As LongLong, FileE_LL As LongLong, FileF_LL As LongLong, FILEG_LL As LongLong, FileH_LL As LongLong
Public Rank1_LL As LongLong, Rank2_LL As LongLong, Rank3_LL As LongLong, Rank4_LL As LongLong, Rank5_LL As LongLong, Rank6_LL As LongLong, Rank7_LL As LongLong, Rank8_LL As LongLong
Public RankLL(8) As LongLong
Public FileLL(8) As LongLong
Public Rank64(63) As Long ' chess board rank 1-8 for 64Bit
Public File64(63) As Long ' chess board file 1-8 for 64Bit
Public ForwardRanksLL(COL_WHITE, 8) As LongLong
Public CenterFilesLL As LongLong

' bitboarda for attacks from a square to a direction
Public SquaresToNorthLL(63)         As LongLong
Public SquaresToNorthEastLL(63)     As LongLong
Public SquaresToEastLL(63)          As LongLong
Public SquaresToSouthEastLL(63)     As LongLong
Public SquaresToSouthLL(63)         As LongLong
Public SquaresToSouthWestLL(63)     As LongLong
Public SquaresToWestLL(63)          As LongLong
Public SquaresToNorthWestLL(63)     As LongLong

Public AttacksForColLL(COL_WHITE) As LongLong
Public AttacksForColPieceLL(COL_WHITE, PT_ALL_PIECES) As LongLong
Public AttacksBy2ForColLL(COL_WHITE) As LongLong

Public LastMove As TMOVE  ' for undo move
Public PrevLastMove As TMOVE


'=============================================================================================================

Public Sub TestChess()
  
  InitEngine
  
  Init64Bit
  InitFileRankLL
  InitAttackHelperBoards
  InitAttacksFromSquareLL
  
'  BoardLL = BoardLL Or PieceMovesLL(WKING, SqToBit(SQ_D4))
'  ShowDebugLL PieceMovesLL(WKING, SqToBit(SQ_E1))
'  BoardLL = BoardLL Or PieceMovesLL(BKING, SqToBit(SQ_H8))
  
'  BoardLL = BoardLL Or PieceMovesLL(WKNIGHT, SqToBit(SQ_D4))
'  BoardLL = BoardLL Or PieceMovesLL(BKNIGHT, SqToBit(SQ_H8))
'  BoardLL = BoardLL Or PieceMovesLL(WKNIGHT, SqToBit(SQ_G6))
   
   ' Pawn attacks
'   PiecesLL(WPAWN) = Bit64ValueLL(SqToBit(SQ_A2)) Or Bit64ValueLL(SqToBit(SQ_H2)) Or Bit64ValueLL(SqToBit(SQ_D4)) Or Bit64ValueLL(SqToBit(SQ_A7)) Or Bit64ValueLL(SqToBit(SQ_H7))
'   BoardLL = PawnAttacksLL(COL_WHITE, PiecesLL(WPAWN))
'   ShowDebugLL PiecesLL(WPAWN)
   
'   PiecesLL(BPAWN) = Bit64ValueLL(SqToBit(SQ_A7)) Or Bit64ValueLL(SqToBit(SQ_H7)) Or Bit64ValueLL(SqToBit(SQ_D5)) Or Bit64ValueLL(SqToBit(SQ_A2)) Or Bit64ValueLL(SqToBit(SQ_H2))
'   ShowDebugLL PiecesLL(BPAWN)
   
'   BoardLL = PawnAttacksLL(COL_BLACK, PiecesLL(BPAWN))
'   Board(SQ_E4) = BPAWN
'
    Dim AttacksLL As LongLong
'    BoardLL = 0
'    BoardLL = Bit64ValueLL(SqToBit(SQ_G7)) Or Bit64ValueLL(SqToBit(SQ_G2))
'    AttacksLL = SliderAttacksFromSq64LL(BQUEEN, SqToBit(SQ_C3), BoardLL)
'    AttacksLL = SliderAttacksFromSq64LL(BQUEEN, SqToBit(SQ_C3), BoardLL)
'    ShowDebugLL AttacksLL

  Erase PieceCnt()
  
  ClearBoard64LL
  SetPiece64LL SQ_H1, WKING: WKingLoc = SQ_H1
  SetPiece64LL SQ_H4, WPAWN: PieceCnt(WPAWN) = PieceCnt(WPAWN) + 1
  SetPiece64LL SQ_G3, WPAWN: PieceCnt(WPAWN) = PieceCnt(WPAWN) + 1
  SetPiece64LL SQ_B2, WPAWN: PieceCnt(WPAWN) = PieceCnt(WPAWN) + 1
  SetPiece64LL SQ_C3, WPAWN: PieceCnt(WPAWN) = PieceCnt(WPAWN) + 1
  SetPiece64LL SQ_F1, WROOK: PieceCnt(WROOK) = PieceCnt(WROOK) + 1
  SetPiece64LL SQ_A1, WROOK: PieceCnt(WROOK) = PieceCnt(WROOK) + 1
  SetPiece64LL SQ_G4, WKNIGHT: PieceCnt(WKNIGHT) = PieceCnt(WKNIGHT) + 1
  SetPiece64LL SQ_E4, WBISHOP: PieceCnt(WBISHOP) = PieceCnt(WBISHOP) + 1
  SetPiece64LL SQ_C2, WQUEEN: PieceCnt(WQUEEN) = PieceCnt(WQUEEN) + 1
  
  SetPiece64LL SQ_H8, BKING: BKingLoc = SQ_H8
  SetPiece64LL SQ_H5, BPAWN: PieceCnt(BPAWN) = PieceCnt(BPAWN) + 1
  SetPiece64LL SQ_G6, BPAWN: PieceCnt(BPAWN) = PieceCnt(BPAWN) + 1
  SetPiece64LL SQ_B7, BPAWN: PieceCnt(BPAWN) = PieceCnt(BPAWN) + 1
  SetPiece64LL SQ_C6, BPAWN: PieceCnt(BPAWN) = PieceCnt(BPAWN) + 1
  SetPiece64LL SQ_F8, BROOK: PieceCnt(BROOK) = PieceCnt(BROOK) + 1
  SetPiece64LL SQ_A8, BROOK: PieceCnt(BROOK) = PieceCnt(BROOK) + 1
  SetPiece64LL SQ_G5, BKNIGHT: PieceCnt(BKNIGHT) = PieceCnt(BKNIGHT) + 1
  SetPiece64LL SQ_E5, BBISHOP: PieceCnt(BBISHOP) = PieceCnt(BBISHOP) + 1
  SetPiece64LL SQ_C7, BQUEEN: PieceCnt(BQUEEN) = PieceCnt(BQUEEN) + 1
'  bWhiteToMove = True
  bWhiteToMove = True
  
  ShowBoardLL
'  GetAttacksLL
  'ShowDebugLL AttacksPiecesLL(WPAWN)
  'ShowDebugLL AttacksForColLL(COL_WHITE)
  'ShowDebugLL AttacksBy2ForColLL(COL_WHITE)

'  ShowDebugLL BoardLL
 ' ShowDebugLL PiecesLL(COL_BLACK, PT_KING)
 ' PiecesLL(COL_BLACK, PT_KING) = Not PiecesLL(COL_BLACK, PT_KING)
 ' ShowDebugLL PiecesLL(COL_BLACK, PT_KING)
  
'  ShowDebugLL PiecesForColLL(COL_WHITE)
'  AttacksLL = SliderAttacksFromSq64LL(WROOK, SqToBit(SQ_B2), BoardLL) Or SliderAttacksFromSq64LL(WROOK, SqToBit(SQ_G4), BoardLL)
    
'  Debug.Print PrintPos
'    Dim NumMoves As Long
'    GenerateMoves 1, COL_WHITE, NumMoves
'
'    Dim MoveCnt As Long
'    For MoveCnt = 0 To NumMoves - 1
'      Debug.Print MoveText(Moves(1, MoveCnt))
'    Next
  
End Sub

Public Sub ClearBoard64LL()
  BoardLL = 0
  Erase PiecesLL
  Erase PiecesForColLL
End Sub

Public Sub SetPiece64LL(Square As Long, Piece As Long)
  Dim SqVal As LongLong
  Board(Square) = Piece
  
  SqVal = Bit64ValueLL(SqToBit(Square))
  BoardLL = BoardLL Or SqVal
  PiecesLL(PieceColor(Piece), PieceType(Piece)) = PiecesLL(PieceColor(Piece), PieceType(Piece)) Or SqVal
  PiecesForColLL(PieceColor(Piece)) = PiecesForColLL(PieceColor(Piece)) Or SqVal
End Sub

Public Sub RemovePiece64LL(Square As Long, Piece As Long)
  Dim SqVal As LongLong
  SqVal = Bit64ValueLL(SqToBit(Square))
  BoardLL = BoardLL And Not SqVal
  PiecesLL(PieceColor(Piece), PieceType(Piece)) = PiecesLL(PieceColor(Piece), PieceType(Piece)) And Not SqVal
  PiecesForColLL(PieceColor(Piece)) = PiecesForColLL(PieceColor(Piece)) And Not SqVal
End Sub



Public Function SliderAttacksFromSq64LL(Piece As Long, Square64 As Long, OccupiedLL As LongLong) As LongLong
  Dim AttackLL As LongLong, BlockerLL As LongLong
  
  SliderAttacksFromSq64LL = 0
  
  '--- positive direction = Lsb64LL  /  negative direction = Rsb64LL
  Select Case Piece ' orthogonal
  Case WQUEEN, BQUEEN, WROOK, BROOK
    AttackLL = SquaresToNorthLL(Square64)
    If AttackLL <> 0 Then
      BlockerLL = OccupiedLL And AttackLL: If BlockerLL <> 0 Then AttackLL = AttackLL Xor SquaresToNorthLL(Lsb64LL(OccupiedLL And AttackLL))
      SliderAttacksFromSq64LL = SliderAttacksFromSq64LL Or AttackLL
    End If
  
    AttackLL = SquaresToEastLL(Square64)
    If AttackLL <> 0 Then
      BlockerLL = OccupiedLL And AttackLL: If BlockerLL <> 0 Then AttackLL = AttackLL Xor SquaresToEastLL(Lsb64LL(OccupiedLL And AttackLL))
      SliderAttacksFromSq64LL = SliderAttacksFromSq64LL Or AttackLL
    End If
    
    AttackLL = SquaresToSouthLL(Square64)
    If AttackLL <> 0 Then
      BlockerLL = OccupiedLL And AttackLL: If BlockerLL <> 0 Then AttackLL = AttackLL Xor SquaresToSouthLL(Rsb64LL(OccupiedLL And AttackLL))
      SliderAttacksFromSq64LL = SliderAttacksFromSq64LL Or AttackLL
    End If
  
    AttackLL = SquaresToWestLL(Square64)
    If AttackLL <> 0 Then
      BlockerLL = OccupiedLL And AttackLL: If BlockerLL <> 0 Then AttackLL = AttackLL Xor SquaresToWestLL(Rsb64LL(OccupiedLL And AttackLL))
      SliderAttacksFromSq64LL = SliderAttacksFromSq64LL Or AttackLL
    End If
  End Select
    
  Select Case Piece ' diagonal
  Case WBISHOP, BBISHOP, WQUEEN, BQUEEN
    AttackLL = SquaresToNorthWestLL(Square64)
    If AttackLL <> 0 Then
      BlockerLL = OccupiedLL And AttackLL: If BlockerLL <> 0 Then AttackLL = AttackLL Xor SquaresToNorthWestLL(Lsb64LL(OccupiedLL And AttackLL))
      SliderAttacksFromSq64LL = SliderAttacksFromSq64LL Or AttackLL
    End If
    
    AttackLL = SquaresToNorthEastLL(Square64)
    If AttackLL <> 0 Then
      BlockerLL = OccupiedLL And AttackLL: If BlockerLL <> 0 Then AttackLL = AttackLL Xor SquaresToNorthEastLL(Lsb64LL(OccupiedLL And AttackLL))
      SliderAttacksFromSq64LL = SliderAttacksFromSq64LL Or AttackLL
    End If
        
    AttackLL = SquaresToSouthWestLL(Square64)
    If AttackLL <> 0 Then
      BlockerLL = OccupiedLL And AttackLL: If BlockerLL <> 0 Then AttackLL = AttackLL Xor SquaresToSouthWestLL(Rsb64LL(OccupiedLL And AttackLL))
      SliderAttacksFromSq64LL = SliderAttacksFromSq64LL Or AttackLL
    End If
    
    AttackLL = SquaresToSouthEastLL(Square64)
    If AttackLL <> 0 Then
      BlockerLL = OccupiedLL And AttackLL: If BlockerLL <> 0 Then AttackLL = AttackLL Xor SquaresToSouthEastLL(Rsb64LL(OccupiedLL And AttackLL))
      SliderAttacksFromSq64LL = SliderAttacksFromSq64LL Or AttackLL
    End If
  End Select
  
End Function


Public Sub InitAttacksFromSquareLL()
  Dim Square As Long, Piece As Long, MoveDir As Long, TargetSq As Long
  Erase PieceMovesLL
  For Square = SQ_A1 To SQ_H8
    If File(Square) >= FILE_A And File(Square) <= FILE_H Then
      For Piece = WPAWN To BKING
      
        Select Case Piece
        '--- King ---
        Case WKING, BKING
          For MoveDir = 0 To 7
            TargetSq = Square + DirectionOffset(MoveDir)
            If Board(TargetSq) <> FRAME Then
              PieceMovesLL(Piece, SqToBit(Square)) = PieceMovesLL(Piece, SqToBit(Square)) Or Bit64ValueLL(SqToBit(TargetSq))
            End If
          Next MoveDir
        '--- Knight ---
        Case WKNIGHT, BKNIGHT
          For MoveDir = 0 To 7
            TargetSq = Square + KnightOffsets(MoveDir)
            If Board(TargetSq) <> FRAME Then
              PieceMovesLL(Piece, SqToBit(Square)) = PieceMovesLL(Piece, SqToBit(Square)) Or Bit64ValueLL(SqToBit(TargetSq))
            End If
          Next MoveDir
        End Select
      Next ' Piece
    End If
  Next ' Square
End Sub


Public Sub GetAttacksLL()
 Dim sq As Long, sqLL As LongLong, Piece As Long, Col As enumColor, PType As enumPieceType, AttackLL As LongLong
 
 For Col = COL_BLACK To COL_WHITE ' set pawn attacks and init other bitboards
   AttackLL = PawnAttacksLL(Col, PiecesLL(Col, PT_PAWN))
   AttacksForColPieceLL(Col, PT_PAWN) = AttackLL
   AttacksBy2ForColLL(Col) = AttackLL ' double attacked?
   AttacksForColLL(Col) = AttackLL
 Next Col
 
 For sq = 0 To 63
   sqLL = Bit64ValueLL(sq)
   If BoardLL And sqLL Then
     For Piece = WKNIGHT To BKING
       Col = PieceColor(Piece): PType = PieceType(Piece)
       If sqLL And PiecesLL(Col, PType) Then
        Select Case PType
        Case PT_KNIGHT, PT_KING
          AttackLL = PieceMovesLL(Piece, sq)
        Case Else ' WBISHOP, WROOK, WQUEEN, BBISHOP, BROOK, BQUEEN
          AttackLL = SliderAttacksFromSq64LL(Piece, sq, BoardLL)
        End Select
        AttacksForColPieceLL(Col, PType) = AttacksForColPieceLL(Col, PType) Or AttackLL
        AttacksBy2ForColLL(Col) = AttacksBy2ForColLL(Col) Or (AttacksForColLL(Col) And AttackLL) ' double attacked?
        AttacksForColLL(Col) = AttacksForColLL(Col) Or AttackLL
        Exit For
       End If
     Next Piece
   End If
 Next sq
 
End Sub

Public Function PawnAttacksLL(ByRef Col As enumColor, op1 As LongLong) As LongLong
  If Col = COL_WHITE Then
    PawnAttacksLL = BitsShiftRightLL(op1, 8) ' Shift Up = 8 * ShiftRight, white pawns never at RANK8. Special case H7 for Bit63 (sign bit)
  ElseIf Col = COL_BLACK Then
    PawnAttacksLL = BitsShiftLeftLL(op1, 8) ' Shift Down = 8 * ShiftLeft, black pawns never at RANK1
  End If
  ' remove FILE A and ShiftLeft +  Remove FILE H and ShiftRight
  PawnAttacksLL = BitsShiftLeftLL(PawnAttacksLL And Not FileA_LL, 1) Or BitsShiftRightLL(PawnAttacksLL And Not FileH_LL, 1)
End Function


Public Sub InitFileRankLL()
 Dim i As Long, sqLL As LongLong, SqBB As Long
  SqBB = 0
  For i = 0 To 119
    SqToBit(i) = -1
    If Board(i) <> FRAME Then
        SqToBit(i) = SqBB ' 0..63
        BitToSq(SqBB) = i
        Rank64(SqBB) = 1 + SqBB \ 8
        File64(SqBB) = 1 + SqBB Mod 8
        
        '--- set ranks
        RankLL(Rank64(SqBB)) = RankLL(Rank64(SqBB)) Or Bit64ValueLL(SqBB)
        Select Case Rank64(SqBB)
        Case 1: Rank1_LL = Rank1_LL Or Bit64ValueLL(SqBB)
        Case 2: Rank2_LL = Rank2_LL Or Bit64ValueLL(SqBB)
        Case 3: Rank3_LL = Rank3_LL Or Bit64ValueLL(SqBB)
        Case 4: Rank4_LL = Rank4_LL Or Bit64ValueLL(SqBB)
        Case 5: Rank5_LL = Rank5_LL Or Bit64ValueLL(SqBB)
        Case 6: Rank6_LL = Rank6_LL Or Bit64ValueLL(SqBB)
        Case 7: Rank7_LL = Rank7_LL Or Bit64ValueLL(SqBB)
        Case 8: Rank8_LL = Rank8_LL Or Bit64ValueLL(SqBB)
        End Select

        '--- set Files
        FileLL(File64(SqBB)) = FileLL(File64(SqBB)) Or Bit64ValueLL(SqBB)
        Select Case File64(SqBB)
        Case 1: FileA_LL = FileA_LL Or Bit64ValueLL(SqBB)
        Case 2: FileB_LL = FileB_LL Or Bit64ValueLL(SqBB)
        Case 3: FileC_LL = FileC_LL Or Bit64ValueLL(SqBB)
        Case 4: FileD_LL = FileD_LL Or Bit64ValueLL(SqBB)
        Case 5: FileE_LL = FileE_LL Or Bit64ValueLL(SqBB)
        Case 6: FileF_LL = FileF_LL Or Bit64ValueLL(SqBB)
        Case 7: FILEG_LL = FILEG_LL Or Bit64ValueLL(SqBB)
        Case 8: FileH_LL = FileH_LL Or Bit64ValueLL(SqBB)
        End Select
        
      'SeLongLong SquareBB(i), SqBB
      'If ColorSq(i) = COL_BLACK Then SeLongLong DarkSquaresBB, SqBB
      '
      
      
      SqBB = SqBB + 1
    End If
  Next i
  CenterFilesLL = FileC_LL Or FileD_LL Or FileE_LL Or FileF_LL
  
  ForwardRanksLL(COL_WHITE, 7) = Rank8_LL
  For i = 6 To 1 Step -1
    ForwardRanksLL(COL_WHITE, i) = ForwardRanksLL(COL_WHITE, i + 1) Or RankLL(i)
  Next
  
  ForwardRanksLL(COL_BLACK, 2) = Rank1_LL
  For i = 3 To 8
    ForwardRanksLL(COL_BLACK, i) = ForwardRanksLL(COL_BLACK, i - 1) Or RankLL(i)
  Next
End Sub

Public Sub InitAttackHelperBoards()
  Dim i As Long, Square As Long
  For i = SQ_A1 To SQ_H8
    If Board(i) <> FRAME Then
      ' Init attack helper bitboards: attacks form square x in one direction
      Square = i + SQ_RIGHT
      Do While Board(Square) <> FRAME
        SquaresToEastLL(SqToBit(i)) = SquaresToEastLL(SqToBit(i)) Or Bit64ValueLL(SqToBit(Square)): Square = Square + SQ_RIGHT
      Loop
      
      Square = i + SQ_DOWN_RIGHT
      Do While Board(Square) <> FRAME
        SquaresToSouthEastLL(SqToBit(i)) = SquaresToSouthEastLL(SqToBit(i)) Or Bit64ValueLL(SqToBit(Square)): Square = Square + SQ_DOWN_RIGHT
      Loop
      
      Square = i + SQ_DOWN
      Do While Board(Square) <> FRAME
        SquaresToSouthLL(SqToBit(i)) = SquaresToSouthLL(SqToBit(i)) Or Bit64ValueLL(SqToBit(Square)): Square = Square + SQ_DOWN
      Loop
      
      Square = i + SQ_DOWN_LEFT
      Do While Board(Square) <> FRAME
        SquaresToSouthWestLL(SqToBit(i)) = SquaresToSouthWestLL(SqToBit(i)) Or Bit64ValueLL(SqToBit(Square)): Square = Square + SQ_DOWN_LEFT
      Loop
      
      Square = i + SQ_LEFT
      Do While Board(Square) <> FRAME
        SquaresToWestLL(SqToBit(i)) = SquaresToWestLL(SqToBit(i)) Or Bit64ValueLL(SqToBit(Square)): Square = Square + SQ_LEFT
      Loop
      
      Square = i + SQ_UP_LEFT
      Do While Board(Square) <> FRAME
        SquaresToNorthWestLL(SqToBit(i)) = SquaresToNorthWestLL(SqToBit(i)) Or Bit64ValueLL(SqToBit(Square)): Square = Square + SQ_UP_LEFT
      Loop
      
      Square = i + SQ_UP
      Do While Board(Square) <> FRAME
        SquaresToNorthLL(SqToBit(i)) = SquaresToNorthLL(SqToBit(i)) Or Bit64ValueLL(SqToBit(Square)): Square = Square + SQ_UP
      Loop
      
      Square = i + SQ_UP_RIGHT
      Do While Board(Square) <> FRAME
        SquaresToNorthEastLL(SqToBit(i)) = SquaresToNorthEastLL(SqToBit(i)) Or Bit64ValueLL(SqToBit(Square)): Square = Square + SQ_UP_RIGHT
      Loop
    End If
  Next

End Sub






'=================== 64 bitboard ==========================================================

Public Sub ShowDebugLL(InBoardLL As LongLong)
 Dim i As Long, s As String
 Debug.Print
 Debug.Print " ------------------"
 s = ""
 For i = 63 To 0 Step -1
   If (i + 1) Mod 8 = 0 And s <> "" Then
     Debug.Print CStr((i + 9) \ 8) & "|" & s & "|"
     s = ""
   End If
   If InBoardLL And Bit64ValueLL(i) Then s = " X" & s Else s = " ." & s
 Next
 Debug.Print CStr((i + 9) \ 8) & "|" & s & "|"
 Debug.Print " ------------------"
 Debug.Print "   A B C D E F G H"
 Debug.Print
End Sub

Public Sub ShowBoardLL()
 Dim i As Long, s As String, Piece As Long, sqLL As LongLong, PieceLL As LongLong
 Debug.Print
 Debug.Print " ------------------"
 s = ""
 
 For i = 63 To 0 Step -1
   If (i + 1) Mod 8 = 0 And s <> "" Then
     Debug.Print CStr((i + 9) \ 8) & "|" & s & "|"
     s = ""
   End If
   sqLL = Bit64ValueLL(i)
   If BoardLL And sqLL Then
     For Piece = WPAWN To BKING
        If sqLL And PiecesLL(PieceColor(Piece), PieceType(Piece)) Then s = " " & Piece2Alpha(Piece) & s: Exit For
     Next Piece
   Else
     s = " ." & s
   End If
 Next
 Debug.Print CStr((i + 9) \ 8) & "|" & s & "|"
 Debug.Print " ------------------"
 Debug.Print "   A B C D E F G H"
 Debug.Print
End Sub

Public Function WriteBB(bb As TBit64) As String
 Dim i As Long, s As String
 WriteBB = " " & vbCrLf & vbCrLf
 WriteBB = WriteBB & " ------------------" & vbCrLf
 s = ""
 For i = 63 To 0 Step -1
   If (i + 1) Mod 8 = 0 And s <> "" Then
     WriteBB = WriteBB & CStr((i + 9) \ 8) & "|" & s & "|" & vbCrLf
     s = ""
   End If
   If IsBitSet64(bb, ByVal i) Then s = " X" & s Else s = " ." & s
 Next
 WriteBB = WriteBB & CStr((i + 9) \ 8) & "|" & s & "|" & vbCrLf
 WriteBB = WriteBB & " ------------------" & vbCrLf
 WriteBB = WriteBB & "   A B C D E F G H" & vbCrLf
 WriteBB = WriteBB & " " & vbCrLf
 
End Function

Public Function BoardShiftUpLL(ByVal InBoardLL As LongLong) As LongLong
  BoardShiftUpLL = BitsShiftRightLL(InBoardLL, 8)
End Function

Public Function BoardShiftDownLL(ByVal InBoardLL As LongLong) As LongLong
  BoardShiftDownLL = BitsShiftLeftLL(InBoardLL, 8)
End Function

#End If

