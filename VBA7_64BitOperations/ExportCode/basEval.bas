Attribute VB_Name = "basEval"

Option Explicit

#If VBA7 And Win64 Then 'Note: Win64 = Office64 bit (not Windows 64 bit)

Const PHASE_MIDGAME               As Long = 128
Const PHASE_ENDGAME               As Long = 0
Public MidGameLimit               As Long
Public EndgameLimit               As Long
Public GamePhase                  As Long

Public PieceTypeValue(PT_KING)    As Long
Public PieceScore(17)             As Long
Public PieceCntCol(COL_WHITE, PT_KING) As Long

Public ScoreVal(PT_KING)            As TScore
'------------------------
'--- Piece square tables
'------------------------
Public Psqt(COL_WHITE, PT_ALL_PIECES, MAX_BOARD)        As TScore
Public PsqVal(1, 16, MAX_BOARD)     As Long ' piece square score for piece: (endgame,piece,square)
'--- Mobility values for pieces
Public MobilityPt(16, 29)           As TScore  ' max 29 for queen
Public MobilityScore(COL_WHITE)     As TScore
Public MobilityAreaLL(COL_WHITE)    As LongLong

'--- SEE data ( static exchange evaluation )
Dim PieceList(0 To 32)                            As Long, Cnt As Long
Dim SwapList(0 To 32)                             As Long, slIndex As Long
Dim Blocker(1 To 32)                              As Long, Block As Long

' Threats
Public ThreatByMinor(6)            As TScore ' Attacker is defended minor (B/N)
Public ThreatByRook(6)             As TScore
Public ThreatByKing                As TScore
Public ThreatByRank As TScore
Public ThreatBySafePawn As TScore
Public ThreatByPawnPush As TScore
Public HangingBonus As TScore
Public OverloadBonus As TScore
Public BackwardMalus As TScore
Public DoubledMalus As TScore
Public IsolatedMalus As TScore
Public Connected(1, 1, 2, 8)      As TScore
Public RookOnFile(1)              As TScore
Public Outpost(PT_BISHOP, 1)      As TScore ' for Knight+Bishop
Public KingProtector As TScore
Public KingRingLL(COL_WHITE) As LongLong

Public KingAttackWeights(6)       As Long
Public KingAttackersCount(COL_WHITE)  As Long, KingAttackersWeight(COL_WHITE) As Long, KingAttacksCount(COL_WHITE) As Long

' Endgame
Public PushClose(8)               As Long
Public PushAway(8)                As Long
Public PushToEdges(MAX_BOARD)     As Long

'--------------------------------------------------------------------------------------------------------------
Public Function EvalMoves(ByVal inCol As enumColor, BestMove As TMOVE, MoveScore As Long) As String
  Dim NumMoves As Long, LegalMoves As Long
  Dim i        As Long, OppCol As enumColor, s As String, OppMove As TMOVE, OppMoveScore As Long, BestOppMove As TMOVE
  
  MoveScore = VALUE_NONE
  BestMove = EmptyMove
  LegalMoves = 0
  
  Ply = Ply + 1
  GenerateMoves Ply, inCol, NumMoves
  If inCol = COL_WHITE Then OppCol = COL_BLACK Else OppCol = COL_WHITE
  
  For i = 0 To NumMoves - 1
    MakeMove Moves(Ply, i)
   ' If Ply = 1 And MoveText(Moves(Ply, i)) = "g8f6" Then Stop
   ' If MoveText(Moves(Ply, i)) = "g7g5" Then Stop
      
    Moves(Ply, i).OrderValue = -VALUE_INFINITE
    Moves(Ply, i).IsLegal = CheckLegal(Moves(Ply, i))
    If Moves(Ply, i).IsLegal Then
      ' ---- test: If MoveText(Moves(Ply, i)) = "f8xf1" Then Stop
      ' Eval always from view of moving side
      Moves(Ply, i).OrderValue = -Eval()
      LegalMoves = LegalMoves + 1
      If Ply = 1 Then
        '>>>>>> recursive call <<<<<<
        s = EvalMoves(OppCol, OppMove, OppMoveScore)

        OppMoveScore = -OppMoveScore ' from view of other side
        Moves(Ply, i).OrderValue = OppMoveScore
        If OppMoveScore <> VALUE_NONE Then
          If OppMoveScore > MoveScore Or MoveScore = VALUE_NONE Then
            MoveScore = OppMoveScore: BestOppMove = OppMove
            BestMove = Moves(Ply, i)
           ' Debug.Print "BEST: " & MoveText(Moves(Ply, i)) & " " & MoveText(BestOppMove) & " " & OppMoveScore
          End If
        End If
       ' Debug.Print MoveText(Moves(Ply, i)) & " " & MoveText(OppMove) & " " & OppMoveScore
      End If
    End If
    UnmakeMove Moves(Ply, i)
  Next
  
  ' sort highest value on top
  SortMovesStable Ply, 0, NumMoves - 1
  
  ' show result
  If Ply = 1 Then
    EvalMoves = "Best: " & MoveText(Moves(Ply, 0)) & " " & MoveText(BestOppMove) & vbCrLf
    EvalMoves = EvalMoves & "Moves___:__Eval" & vbCrLf
    For i = 0 To NumMoves - 1
      EvalMoves = EvalMoves & Left$(MoveText(Moves(Ply, i)) & Space(8), 8) & ":" & Right$(Space(6) & CStr(EvalTo100(Moves(Ply, i).OrderValue)), 6) & vbCrLf
    Next i
    EvalMoves = EvalMoves & "==============="
  Else
    BestMove = Moves(Ply, 0): MoveScore = Moves(Ply, 0).OrderValue
  End If
  
  If LegalMoves = 0 Then
    If InCheck() Then
      MoveScore = -(99999 - Ply)
    Else
      MoveScore = 0
    End If
  End If
  
  Ply = Ply - 1
End Function


Public Function Eval() As Long
 Dim bbLL As LongLong, AttackLL As LongLong, Score As TScore, Square As Long, Piece As Long, Pt As Long, Col As enumColor, SqVal As LongLong
 Dim MaterialScore As TScore, PositionScore As TScore, Us As enumColor, Them As enumColor
 Dim WPawnCnt As Long, BPawnCnt As Long
 Dim Material As Long, WMaterial As Long, BMaterial As Long, WNonPawnMaterial As Long, BNonPawnMaterial As Long, NonPawnMaterial As Long
 
 InitEval
 
 MaterialScore.MG = 0: MaterialScore.EG = 0
 PositionScore.MG = 0: PositionScore.EG = 0
 Erase MobilityScore()
 
 ' init bitboards
 BoardLL = 0: Erase PiecesLL(): Erase PiecesForColLL()
 
 ' fill main bitboards, count pieces
 Erase PieceCntCol
 For Square = SQ_A1 To SQ_H8
    Piece = Board(Square)
    Select Case Piece
    Case NO_PIECE, FRAME, WEP_PIECE, BEP_PIECE ' ignore
    Case Else
      Col = PieceColor(Piece): Pt = PieceType(Piece)
      PieceCntCol(Col, Pt) = PieceCntCol(Col, Pt) + 1
      If Col = COL_WHITE Then
       AddScore PositionScore, Psqt(Col, Pt, Square)
       ' Debug.Print "WHITE " & PT & " " & Square & " > " & Psqt(Col, PT, Square).MG
      Else
       MinusScore PositionScore, Psqt(Col, Pt, Square)
       ' Debug.Print "BLACK " & PT & " " & Square & " > " & Psqt(Col, PT, Square).MG
      End If
      
      '--- Fill pieces bitboards
      SqVal = Bit64ValueLL(SqToBit(Square))
      BoardLL = BoardLL Or SqVal
      PiecesLL(Col, Pt) = PiecesLL(Col, Pt) Or SqVal
      PiecesForColLL(Col) = PiecesForColLL(Col) Or SqVal
    End Select
 Next
 
  WPawnCnt = PieceCntCol(COL_WHITE, PT_PAWN): BPawnCnt = PieceCntCol(COL_BLACK, PT_PAWN)
  WNonPawnMaterial = 0: BNonPawnMaterial = 0
 ' WKingFile = File(WKingLoc): BKingFile = File(BKingLoc)
  
  For Pt = PT_PAWN To PT_QUEEN
     MaterialScore.MG = MaterialScore.MG + ScoreVal(Pt).MG * (PieceCntCol(COL_WHITE, Pt) - PieceCntCol(COL_BLACK, Pt))
     MaterialScore.EG = MaterialScore.EG + ScoreVal(Pt).EG * (PieceCntCol(COL_WHITE, Pt) - PieceCntCol(COL_BLACK, Pt))
     
     If Pt <> PT_PAWN Then
       WNonPawnMaterial = WNonPawnMaterial + PieceCntCol(COL_WHITE, Pt) * PieceTypeValue(Pt)
       BNonPawnMaterial = BNonPawnMaterial + PieceCntCol(COL_BLACK, Pt) * PieceTypeValue(Pt)
     End If
  Next Pt
  
  WMaterial = WNonPawnMaterial + WPawnCnt * PieceTypeValue(PT_PAWN)
  BMaterial = BNonPawnMaterial + BPawnCnt * PieceTypeValue(PT_PAWN)
  NonPawnMaterial = WNonPawnMaterial + BNonPawnMaterial
  Material = WMaterial - BMaterial
  
   'If Board(SQ_D4) = WQUEEN And Board(SQ_C6) = BKNIGHT Then Stop
  
  ' Init pawn attacks and init other attack bitboards
  Erase AttacksForColPieceLL()
  For Col = COL_BLACK To COL_WHITE ' set pawn attacks and init other bitboards
    AttackLL = PawnAttacksLL(Col, PiecesLL(Col, PT_PAWN))
    AttacksForColPieceLL(Col, PT_PAWN) = AttackLL
    AttacksBy2ForColLL(Col) = AttackLL ' double attacked?
    AttacksForColLL(Col) = AttackLL
  Next Col
  

  ' Init Mobility relevant squares
  For Col = COL_BLACK To COL_WHITE
    bbLL = 0
    ' Find our pawns that are blocked or on the first two ranks
    If Col = COL_WHITE Then
      Us = COL_WHITE: Them = COL_BLACK: bbLL = PiecesLL(Us, PT_PAWN) And (BoardShiftDownLL(BoardLL) Or (Rank2_LL Or Rank3_LL)) ' shift pieces down to pawn rank
    Else
      Us = COL_BLACK: Them = COL_WHITE:  bbLL = PiecesLL(Us, PT_PAWN) And (BoardShiftUpLL(BoardLL) Or (Rank6_LL Or Rank7_LL)) ' shift pieces down to pawn rank
    End If
    
    ' Squares occupied by those pawns, by our king or queen or controlled by
    ' enemy pawns are excluded from the mobility area.
    bbLL = bbLL Or PiecesLL(Us, PT_KING) Or PiecesLL(Us, PT_QUEEN)
    bbLL = bbLL Or AttacksForColPieceLL(Them, PT_PAWN)
    MobilityAreaLL(Col) = Not bbLL  ' use the opposite of these square for pieces mobility scoring
  Next
 'If Ply = 2 And Board(SQ_D4) = WPAWN And Board(SQ_A5) = BPAWN Then Stop
 Score.MG = MaterialScore.MG + PositionScore.MG: Score.EG = MaterialScore.EG + PositionScore.EG
  
 NonPawnMaterial = GetMax(EndgameLimit, GetMin(NonPawnMaterial, MidGameLimit))
 GamePhase = (((NonPawnMaterial - EndgameLimit) * PHASE_MIDGAME) / (MidGameLimit - EndgameLimit))
 
 '------------------------------------------------------------------------------------------------------
 ' pieces eval
 InitKingSafety (COL_WHITE)
 InitKingSafety (COL_BLACK)
   
  For Pt = PT_KNIGHT To PT_KING  ' also adds attacks info for non pawn pieces
     AddScore Score, EvalNonPawnPieces(COL_WHITE, Pt)
     MinusScore Score, EvalNonPawnPieces(COL_BLACK, Pt)
  Next Pt
  
  '--- Endgame KXK ?  i.e. KQ-K or KR-K or KBN-K
  '--- here becasue all attacks needed
  If WNonPawnMaterial >= PieceTypeValue(PT_ROOK) And BMaterial = 0 Then
    Eval = Endgame_KXK(COL_WHITE, WNonPawnMaterial)
     If Eval = 0 Then Stop
    Exit Function  '<<<<<<<<<< EXIT
  ElseIf BNonPawnMaterial >= PieceTypeValue(PT_ROOK) And WMaterial = 0 Then
    Eval = Endgame_KXK(COL_BLACK, BNonPawnMaterial)
    Exit Function  '<<<<<<<<<< EXIT
  End If
  'Debug.Print "Material+Position: " & Score.MG
  

  ' Eval pawns
  AddScore Score, EvalPawns(COL_WHITE)
  MinusScore Score, EvalPawns(COL_BLACK)
  ' Debug.Print "Pawns: " & Score.MG
 
  ' King safety
  AddScore Score, EvalKingSafety(COL_WHITE)
  MinusScore Score, EvalKingSafety(COL_BLACK)
  'Debug.Print "King: " & Score.MG

  '  Add mobility
  AddScore Score, MobilityScore(COL_WHITE)
  MinusScore Score, MobilityScore(COL_BLACK)
  ' Debug.Print "Mobility: " & Score.MG

  ' Add threats
  AddScore Score, EvalThreats(COL_WHITE)
  MinusScore Score, EvalThreats(COL_BLACK)
  'Debug.Print "Threats: " & Score.MG

  ' Add Space
  If NonPawnMaterial > 12222 Then
    AddScore Score, EvalSpace(COL_WHITE)
    MinusScore Score, EvalSpace(COL_BLACK)
   ' Debug.Print "Space: " & Score.MG
  End If

 '-----------------------------------------------------------------------------------------------------
 
 '--- Scale score to game phase
 Eval = (Score.MG * GamePhase + Score.EG * CLng(PHASE_MIDGAME - GamePhase)) \ PHASE_MIDGAME
 If Not bWhiteToMove Then Eval = -Eval ' always from view of moving side
 
End Function

Public Function EvalPawns(ByVal Col As enumColor) As TScore
  Dim SqBit As Long, Square As Long, MobCnt As Long, Us As enumColor, Them As enumColor, RankUp As Long
  Dim DoubledLL As LongLong, SupportedLL As LongLong, PhalanxLL As LongLong, OurPawnsLL As LongLong, TheirPawnsLL As LongLong, DoubledMaskLL As LongLong, OpposedLL As LongLong
  Dim NeighboursLL As LongLong, IsolatedLL As LongLong, PawnFile As Long, PawnRank As Long
  Dim PieceLL As LongLong, Score As TScore
  
  Score.MG = 0: Score.EG = 0
  OurPawnsLL = PiecesLL(Col, PT_PAWN)
  
  If Col = COL_WHITE Then
    Us = COL_WHITE: Them = COL_BLACK
    DoubledMaskLL = BoardShiftDownLL(OurPawnsLL): RankUp = 1
  Else
    Us = COL_BLACK: Them = COL_WHITE
    DoubledMaskLL = BoardShiftUpLL(OurPawnsLL): RankUp = -1
  End If
  TheirPawnsLL = PiecesLL(Them, PT_PAWN)
  
  
  PieceLL = OurPawnsLL
  Do While PieceLL <> 0
    SqBit = PopLsb64LL(PieceLL) ' looks for position of first bit set and removes it
    Square = BitToSq(SqBit): PawnFile = File(Square): PawnRank = Rank(Square)
    
    OpposedLL = TheirPawnsLL And (ForwardRanksLL(Us, PawnRank) And FileLL(PawnFile))
    NeighboursLL = OurPawnsLL And AdjacentFilesLL(PawnFile)
    PhalanxLL = NeighboursLL And RankLL(PawnRank)
    SupportedLL = NeighboursLL And RankLL(PawnRank - RankUp)
    
    If SupportedLL Or PhalanxLL Then
      AddScore Score, Connected(Abs(CBool(OpposedLL)), Abs(CBool(PhalanxLL)), PopCnt64LL(SupportedLL), RelativeRank(Us, Square))
    ElseIf Not NeighboursLL Then
      MinusScore Score, IsolatedMalus
    End If
    
    ' Doubled pawns
    DoubledLL = DoubledMaskLL And Bit64ValueLL(SqBit)
    If DoubledLL And Not SupportedLL Then
      MinusScore Score, DoubledMalus
    End If
  Loop
  EvalPawns = Score
End Function


Public Function EvalNonPawnPieces(ByVal Col As enumColor, PieceType As enumPieceType) As TScore
  ' eval all pieces except pawns for color and piece type
  Dim bbPieceTypeLL As LongLong, AttackLL As LongLong, KingAttackLL As LongLong
  Dim SqBit As Long, Square As Long, MobCnt As Long, Us As enumColor, Them As enumColor, Piece As Long, OwnKingLoc As Long, OppKingLoc As Long
  
  EvalNonPawnPieces.MG = 0: EvalNonPawnPieces.EG = 0
  If Col = COL_WHITE Then
    Us = COL_WHITE: Them = COL_BLACK: OwnKingLoc = WKingLoc: OppKingLoc = BKingLoc
  Else
    Us = COL_BLACK: Them = COL_WHITE: OwnKingLoc = BKingLoc: OppKingLoc = WKingLoc
  End If
  
  '---  for all pieces of this color and type
  bbPieceTypeLL = PiecesLL(Col, PieceType)
  Do While bbPieceTypeLL <> 0
    SqBit = PopLsb64LL(bbPieceTypeLL) ' looks for position of first bit set and removes it
    Square = BitToSq(SqBit): Piece = Board(Square)
    
    Select Case PieceType
    Case PT_KNIGHT, PT_KING
      AttackLL = PieceMovesLL(Piece, SqBit)   ' same for Black
    Case Else ' WBISHOP, WROOK, WQUEEN, BBISHOP, BROOK, BQUEEN
      AttackLL = SliderAttacksFromSq64LL(Piece, SqBit, BoardLL)
    End Select
    AttacksForColPieceLL(Col, PieceType) = AttacksForColPieceLL(Col, PieceType) Or AttackLL
    AttacksBy2ForColLL(Col) = AttacksBy2ForColLL(Col) Or (AttacksForColLL(Col) And AttackLL) ' double attacked?
    AttacksForColLL(Col) = AttacksForColLL(Col) Or AttackLL
    
    ' Mobility
    MobCnt = PopCnt64LL(AttackLL And MobilityAreaLL(Us))
    AddScore MobilityScore(Us), MobilityPt(PieceType, MobCnt)
    
    ' update king attacks
    KingAttackLL = AttackLL And KingRingLL(Them)
    If KingAttackLL <> 0 Then
      If PopCnt64LL(KingAttackLL) > 0 Then
        KingAttackersCount(Us) = KingAttackersCount(Us) + 1
        KingAttackersWeight(Us) = KingAttackersWeight(Us) + KingAttackWeights(PieceType)
        KingAttackLL = AttackLL And PieceMovesLL(Piece, SqToBit(OppKingLoc))
        If KingAttackLL <> 0 Then KingAttacksCount(Us) = KingAttacksCount(Us) + PopCnt64LL(KingAttackLL)
      End If
    End If
    
    'Penalty if the piece is far from the king
    MinusScoreWithFactor EvalNonPawnPieces, KingProtector, MaxDistance(Square, OwnKingLoc)
    
  Loop
   
End Function

Public Function EvalThreats(ByVal Col As enumColor) As TScore
  Dim Score As TScore, WeakLL As LongLong, DefendedLL As LongLong, NonPawnEnemiesLL As LongLong, StronglyProtectedLL As LongLong, SafeLL As LongLong
  Dim SqBit As Long, Square As Long, Us As enumColor, Them As enumColor, Piece As Long, bbLL As LongLong
  
  Score.MG = 0: Score.EG = 0
  If Col = COL_WHITE Then
    Us = COL_WHITE: Them = COL_BLACK
  Else
    Us = COL_BLACK: Them = COL_WHITE
  End If
  
  ' Non-pawn enemies
  NonPawnEnemiesLL = PiecesForColLL(Them) And Not PiecesLL(Them, PT_PAWN)
  
  ' Squares strongly protected by the enemy, either because they defend the
  ' square with a pawn, or because they defend the square twice and we don't.
  StronglyProtectedLL = AttacksForColPieceLL(Them, PT_PAWN) Or (AttacksBy2ForColLL(Them) And Not AttacksBy2ForColLL(Us))
  
  ' Non-pawn enemies, strongly protected
  DefendedLL = NonPawnEnemiesLL And StronglyProtectedLL
  
  ' Enemies not strongly protected and under our attack
  WeakLL = PiecesForColLL(Them) And Not StronglyProtectedLL And AttacksForColLL(Us)
  
  ' Safe or protected squares
  SafeLL = (Not AttacksForColLL(Them)) Or AttacksForColLL(Us)
  
  ' Bonus according to the kind of attacking pieces
  If DefendedLL Or WeakLL Then
  
    ' Attacks by Knight or Bishop
    bbLL = (DefendedLL Or WeakLL) And (AttacksForColPieceLL(Us, PT_KNIGHT) Or AttacksForColPieceLL(Us, PT_BISHOP))
    Do While bbLL
      SqBit = PopLsb64LL(bbLL) ' looks for position of first bit set and removes it
      Square = BitToSq(SqBit): Piece = Board(Square)
      AddScore Score, ThreatByMinor(PieceType(Piece))
      If PieceType(Piece) <> PT_PAWN Then AddScoreWithFactor Score, ThreatByRank, RelativeRank(Them, Square)
    Loop
    
    ' Attacks by Rook
    bbLL = WeakLL And AttacksForColPieceLL(Us, PT_ROOK)
    Do While bbLL
      SqBit = PopLsb64LL(bbLL) ' looks for position of first bit set and removes it
      Square = BitToSq(SqBit): Piece = Board(Square)
      AddScore Score, ThreatByRook(PieceType(Piece))
      If PieceType(Piece) <> PT_PAWN Then AddScoreWithFactor Score, ThreatByRank, RelativeRank(Them, Square)
    Loop
    
    ' Attacks by King
    If WeakLL And AttacksForColPieceLL(Us, PT_KING) Then AddScore Score, ThreatByKing
    
    ' hanging opponent pieces
    bbLL = WeakLL And Not AttacksForColLL(Them)
    If bbLL Then AddScoreWithFactor Score, HangingBonus, PopCnt64LL(bbLL)
    
    bbLL = WeakLL And NonPawnEnemiesLL And AttacksForColLL(Them)
    If bbLL Then AddScoreWithFactor Score, OverloadBonus, PopCnt64LL(bbLL)
    
  End If
  
  ' Find squares where our pawns can push on the next move
  If Us = COL_WHITE Then
    bbLL = BoardShiftUpLL(PiecesLL(Us, PT_PAWN)) And Not BoardLL
    bbLL = bbLL Or (BoardShiftUpLL(bbLL And Rank3_LL) And Not BoardLL) ' double pawn step
  Else
    bbLL = BoardShiftDownLL(PiecesLL(Us, PT_PAWN)) And Not BoardLL
    bbLL = bbLL Or (BoardShiftDownLL(bbLL And Rank6_LL) And Not BoardLL) ' double pawn step
  End If
  ' Keep only the squares which are relatively safe
  bbLL = bbLL And Not AttacksForColPieceLL(Them, PT_PAWN) And SafeLL
  ' Bonus for safe pawn threats on the next move
  bbLL = PawnAttacksLL(Us, bbLL) And PiecesForColLL(Them)
  If bbLL Then AddScoreWithFactor Score, ThreatByPawnPush, PopCnt64LL(bbLL)
  
  ' Our safe or protected pawns > attacks on non pawn enemies
  bbLL = PiecesLL(Us, PT_PAWN) And SafeLL
  bbLL = PawnAttacksLL(Us, bbLL) And NonPawnEnemiesLL
  If bbLL Then AddScoreWithFactor Score, ThreatBySafePawn, PopCnt64LL(bbLL)
  
  ' Own pieces hanging?
  If Us <> SideToMove Then
    Dim TmpMove As TMOVE, SEEval As Long, MinSEEval As Long
    bbLL = PiecesForColLL(Us) And AttacksForColLL(Them)
    MinSEEval = 0
    Do While bbLL <> 0 ' examine all attacked own pieces
      SqBit = PopLsb64LL(bbLL): Square = BitToSq(SqBit)
      
      TmpMove.From = Square: TmpMove.Target = Square: TmpMove.Piece = Board(Square): TmpMove.Captured = NO_PIECE: TmpMove.SeeValue = 0
      ' Move back to old square, were we in danger there?
      SEEval = GetSEE(TmpMove)
      If SEEval < MinSEEval Then MinSEEval = SEEval
    Loop
    ' if hanging piece found, reduced thread value and add negative SEE value
    If MinSEEval < 0 Then
      Score.MG = Score.MG + MinSEEval \ 2: Score.EG = Score.EG + MinSEEval \ 2
    End If
  End If
 
 
  EvalThreats = Score
End Function

Public Sub InitKingSafety(ByVal Col As enumColor)
  ' King safety
  Dim Score As TScore, Us As enumColor, Them As enumColor, OwnKingLoc As Long, Piece As Long, AttackLL As LongLong
  If Col = COL_WHITE Then
    Us = COL_WHITE: Them = COL_BLACK: OwnKingLoc = WKingLoc: Piece = WKING
  Else
    Us = COL_BLACK: Them = COL_WHITE: OwnKingLoc = BKingLoc: Piece = BKING
  End If
  KingAttackersCount(Them) = 0: KingAttackersWeight(Them) = 0: KingAttacksCount(Them) = 0
  
  ' define safety zone arround king, one square - plus one rank
  ' if FILE A or H then move king center to mid of board
  If File(OwnKingLoc) = FILE_A Then
    OwnKingLoc = OwnKingLoc + SQ_RIGHT
  ElseIf File(OwnKingLoc) = FILE_H Then
    OwnKingLoc = OwnKingLoc + SQ_LEFT
  End If
  
  KingRingLL(Col) = PieceMovesLL(Piece, SqToBit(OwnKingLoc))
  
  If Col = COL_WHITE Then
    If Rank(OwnKingLoc) = 1 Then KingRingLL(Us) = KingRingLL(Us) Or BoardShiftUpLL(KingRingLL(Col))
  Else
    If Rank(OwnKingLoc) = 8 Then KingRingLL(Us) = KingRingLL(Us) Or BoardShiftDownLL(KingRingLL(Col))
  End If
  
  ' pawn attacks to king ring
  AttackLL = KingRingLL(Us) And AttacksForColPieceLL(Them, PT_PAWN)
  If AttackLL <> 0 Then KingAttackersCount(Them) = PopCnt64LL(AttackLL)
  
End Sub

Public Function EvalKingSafety(ByVal Col As enumColor) As TScore
  '
  ' King safety eval
  '
  Const QUEEN_SAFE_CHECK  As Long = 780
  Const ROOK_SAFE_CHECK   As Long = 880
  Const BISHOP_SAFE_CHECK As Long = 435
  Const KNIGHT_SAFE_CHECK As Long = 790
  
  Dim Score As TScore, Us As enumColor, Them As enumColor, OwnKingLoc As Long, Piece As Long, AttackLL As LongLong
  Dim KingDanger As Long, WeakLL As LongLong, SafeLL As LongLong, UnSafeChecksLL As LongLong, RookMovesLL As LongLong, BishopMovesLL As LongLong
  
  If Col = COL_WHITE Then
    Us = COL_WHITE: Them = COL_BLACK: OwnKingLoc = WKingLoc: Piece = WKING
  Else
    Us = COL_BLACK: Them = COL_WHITE: OwnKingLoc = BKingLoc: Piece = BKING
  End If
  
  UnSafeChecksLL = 0: KingDanger = 0
  
  ' Attacked squares defended at most once by our queen or king
  WeakLL = AttacksForColLL(Them) And Not AttacksBy2ForColLL(Us) And _
           (Not AttacksForColLL(Us) Or AttacksForColPieceLL(Us, PT_KING) Or AttacksForColPieceLL(Us, PT_QUEEN))
  
  ' Analyse the safe enemy's checks which are possible on next move
  SafeLL = Not PiecesForColLL(Them)
  SafeLL = SafeLL And (Not AttacksForColLL(Us) Or (WeakLL And AttacksBy2ForColLL(Them)))
  
  ' Attacks to king, own queen removed because not good as blocker
  RookMovesLL = SliderAttacksFromSq64LL(WROOK, SqToBit(OwnKingLoc), (BoardLL And Not PiecesLL(Us, PT_QUEEN)))
  BishopMovesLL = SliderAttacksFromSq64LL(WBISHOP, SqToBit(OwnKingLoc), (BoardLL And Not PiecesLL(Us, PT_QUEEN)))
  
  ' Enemy queen safe checks
  If (RookMovesLL Or BishopMovesLL) And AttacksForColPieceLL(Them, PT_QUEEN) And SafeLL And Not AttacksForColPieceLL(Us, PT_QUEEN) Then
    KingDanger = KingDanger + QUEEN_SAFE_CHECK
  End If
  
  ' Checking moves for rook / bishop
  RookMovesLL = RookMovesLL And AttacksForColPieceLL(Them, PT_ROOK)
  BishopMovesLL = BishopMovesLL And AttacksForColPieceLL(Them, PT_BISHOP)
  
  ' Enemy rook checks
  If RookMovesLL And SafeLL Then
    KingDanger = KingDanger + ROOK_SAFE_CHECK
  Else
    UnSafeChecksLL = UnSafeChecksLL Or RookMovesLL
  End If
  
  ' Enemy bishop checks
  If BishopMovesLL And SafeLL Then
    KingDanger = KingDanger + BISHOP_SAFE_CHECK
  Else
    UnSafeChecksLL = UnSafeChecksLL Or BishopMovesLL
  End If
  
  ' Enemy knight checks
  AttackLL = PieceMovesLL(WKNIGHT, SqToBit(OwnKingLoc)) And AttacksForColPieceLL(Them, PT_KNIGHT)
  If AttackLL And SafeLL Then
    KingDanger = KingDanger + KNIGHT_SAFE_CHECK
  Else
    UnSafeChecksLL = UnSafeChecksLL Or AttackLL
  End If
  
  ' Unsafe checks only for mobility area
  UnSafeChecksLL = UnSafeChecksLL And MobilityAreaLL(Them)
  
  '--- DEBUG ---
  ' ShowBoardLL
  ' ShowDebugLL AttacksForColPieceLL(Them, PT_KNIGHT)
  '-------------
  
  '--------------------------
  
  KingDanger = KingDanger + KingAttackersCount(Them) * KingAttackersWeight(Them) _
                     + 69 * KingAttacksCount(Them) _
                     + 185 * PopCnt64LL(KingRingLL(Us) And WeakLL) _
                     + 150 * PopCnt64LL(UnSafeChecksLL) _
                     - 873 * Abs(PieceCntCol(Them, PT_QUEEN) = 0) _
                     - 30
                     
  If KingDanger > 0 Then
    Dim MobilityDanger As Long
    MobilityDanger = MobilityScore(Them).MG - MobilityScore(Us).MG
    KingDanger = GetMax(0, KingDanger + MobilityDanger)
    Score.MG = Score.MG - (KingDanger * KingDanger \ 4096)
    Score.EG = Score.EG - KingDanger \ 16
  End If
    
  EvalKingSafety = Score
End Function
  
Public Function EvalSpace(ByVal Col As enumColor) As TScore
  ' Space evaluation is a simple bonus based on the number of safe squares
  ' available for minor pieces on the central four files on ranks 2--4.
  Dim Score As TScore, Us As enumColor, Them As enumColor, SafeLL As LongLong, SpaceMaskLL As LongLong, BehindLL As LongLong
  Dim Bonus As Long, Weight As Long
   
  If Col = COL_WHITE Then
    Us = COL_WHITE: Them = COL_BLACK
    SpaceMaskLL = CenterFilesLL And (Rank2_LL Or Rank3_LL Or Rank4_LL)
  Else
    Us = COL_BLACK: Them = COL_WHITE
    SpaceMaskLL = CenterFilesLL And (Rank5_LL Or Rank6_LL Or Rank7_LL)
  End If
   
  ' Find the available squares for our pieces inside the area defined by SpaceMask
  SafeLL = SpaceMaskLL And Not PiecesLL(Us, PT_PAWN) And Not AttacksForColPieceLL(Them, PT_PAWN)
   
  ' Find all squares which are at most three squares behind some friendly pawn
  BehindLL = PiecesLL(Us, PT_PAWN)
  If Us = COL_WHITE Then  ' Shift one rank
    BehindLL = BehindLL Or BoardShiftDownLL(BehindLL)
  Else
    BehindLL = BehindLL Or BoardShiftUpLL(BehindLL)
  End If
  If Us = COL_WHITE Then ' Shift second rank
    BehindLL = BehindLL Or BoardShiftDownLL(BehindLL)
  Else
    BehindLL = BehindLL Or BoardShiftUpLL(BehindLL)
  End If
  Bonus = PopCnt64LL(SafeLL) + PopCnt64LL(BehindLL And SafeLL)
  If Bonus Then
    Weight = PopCnt64LL(PiecesForColLL(Us)) - 2 * OpenFiles()
    If Weight Then Score.MG = Bonus * Weight * Weight \ 16
  End If
  EvalSpace = Score
End Function



Public Sub InitEval()
  Static InitDone As Boolean
  
  If InitDone Then
    Exit Sub
  Else
    InitDone = True
  End If
  
  MidGameLimit = 15258 ' total material for midgame phase
  EndgameLimit = 3915  ' total material for endgame phase
  
  ScoreVal(PT_PAWN).MG = 128
  ScoreVal(PT_PAWN).EG = 213
  ScoreVal(PT_KNIGHT).MG = 782
  ScoreVal(PT_KNIGHT).EG = 865
  ScoreVal(PT_BISHOP).MG = 830
  ScoreVal(PT_BISHOP).EG = 918
  ScoreVal(PT_ROOK).MG = 1289
  ScoreVal(PT_ROOK).EG = 1378
  ScoreVal(PT_QUEEN).MG = 2529
  ScoreVal(PT_QUEEN).EG = 2687
  
  PieceTypeValue(PT_PAWN) = ScoreVal(PT_PAWN).MG
  PieceTypeValue(PT_KNIGHT) = ScoreVal(PT_KNIGHT).MG
  PieceTypeValue(PT_BISHOP) = ScoreVal(PT_BISHOP).MG
  PieceTypeValue(PT_ROOK) = ScoreVal(PT_ROOK).MG
  PieceTypeValue(PT_QUEEN) = ScoreVal(PT_QUEEN).MG
  PieceTypeValue(PT_KING) = 0
  
  PieceScore(FRAME) = 0
  PieceScore(WPAWN) = PieceTypeValue(PT_PAWN): PieceScore(BPAWN) = -PieceScore(WPAWN)
  PieceScore(WKNIGHT) = PieceTypeValue(PT_KNIGHT): PieceScore(BKNIGHT) = -PieceScore(WKNIGHT)
  PieceScore(WBISHOP) = PieceTypeValue(PT_BISHOP): PieceScore(BBISHOP) = -PieceScore(WBISHOP)
  PieceScore(WROOK) = PieceTypeValue(PT_ROOK): PieceScore(BROOK) = -PieceScore(WROOK)
  PieceScore(WQUEEN) = PieceTypeValue(PT_QUEEN): PieceScore(BQUEEN) = -PieceScore(WQUEEN)
  PieceScore(WKING) = 5000: PieceScore(BKING) = -PieceScore(WKING)
  PieceScore(13) = 0: PieceScore(14) = 0
  PieceScore(WEP_PIECE) = PieceTypeValue(PT_PAWN): PieceScore(BEP_PIECE) = -PieceTypeValue(PT_PAWN)
  
  ' ( FILE A-D: Pairs MG,EG :  A(MG,EG),B(MG,EG),...
  '--- Pawn piece square table
  PSQT64 PT_PAWN, 0, 0, 0, 0, 0, 0, 0, 0, -11, 7, 6, -4, 7, 8, 3, -2, -18, -4, -2, -5, 19, 5, 24, 4, -17, 3, -9, 3, 20, -8, 35, -3, -6, 8, 5, 9, 3, 7, 21, -6, -6, 8, -8, -5, -6, 2, -2, 4, -4, 3, 20, -9, -8, 1, -4, 18, 0, 0, 0, 0, 0, 0, 0, 0
  '--- Knight piece square table
  PSQT64 PT_KNIGHT, -161, -105, -96, -82, -80, -46, -73, -14, -83, -69, -43, -54, -21, -17, -10, 9, -71, -50, -22, -39, 0, -7, 9, 28, -25, -41, 18, -25, 43, 6, 47, 38, -26, -46, 16, -25, 38, 3, 50, 40, -11, -54, 37, -38, 56, -7, 65, 27, -63, -65, -19, -50, 5, -24, 14, 13, -195, -109, -67, -89, -42, -50, -29, -13
  '--- Bishop piece square table
  PSQT64 PT_BISHOP, -44, -58, -13, -31, -25, -37, -34, -19, -20, -34, 20, -9, 12, -14, 1, 4, -9, -23, 27, 0, 21, -3, 11, 16, -11, -26, 28, -3, 21, -5, 10, 16, -11, -26, 27, -4, 16, -7, 9, 14, -17, -24, 16, -2, 12, 0, 2, 13, -23, -34, 17, -10, 6, -12, -2, 6, -35, -55, -11, -32, -19, -36, -29, -17
  '--- Rook piece square table
  PSQT64 PT_ROOK, -25, 0, -16, 0, -16, 0, -9, 0, -21, 0, -8, 0, -3, 0, 0, 0, -21, 0, -9, 0, -4, 0, 2, 0, -22, 0, -6, 0, -1, 0, 2, 0, -22, 0, -7, 0, 0, 0, 1, 0, -21, 0, -7, 0, 0, 0, 2, 0, -12, 0, 4, 0, 8, 0, 12, 0, -23, 0, -15, 0, -11, 0, -5, 0
  '--- Queen piece square table
  PSQT64 PT_QUEEN, 0, -71, -4, -56, -3, -42, -1, -29, -4, -56, 6, -30, 9, -21, 8, -5, -2, -39, 6, -17, 9, -8, 9, 5, -1, -29, 8, -5, 10, 9, 7, 19, -3, -27, 9, -5, 8, 10, 7, 21, -2, -40, 6, -16, 8, -10, 10, 3, -2, -55, 7, -30, 7, -21, 6, -6, -1, -74, -4, -55, -1, -43, 0, -30
  '--- King piece square table
  PSQT64 PT_KING, 267, 0, 320, 48, 270, 75, 195, 84, 264, 43, 304, 92, 238, 143, 180, 132, 200, 83, 245, 138, 176, 167, 110, 165, 177, 106, 185, 169, 148, 169, 110, 179, 149, 108, 177, 163, 115, 200, 66, 203, 118, 95, 159, 155, 84, 176, 41, 174, 87, 50, 128, 99, 63, 122, 20, 139, 63, 9, 88, 55, 47, 80, 0, 90
  FillPieceSquareVal
  
  '---  Mobility bonus for number of attacked squares not occupied by friendly pieces (pairs: MG,EG, MG,EG)
  '---  indexed by piece type and number of attacked squares in the mobility area.
  ' Knights
  ReadMobScoreArr PT_KNIGHT, -75, -76, -56, -54, -9, -26, -2, -10, 6, 5, 15, 11, 22, 26, 30, 28, 36, 29
  ' Bishops
  ReadMobScoreArr PT_BISHOP, -48, -58, -21, -19, 16, -2, 26, 12, 37, 22, 51, 42, 54, 54, 63, 58, 65, 63, 71, 70, 79, 74, 81, 86, 92, 90, 97, 94
  ' Rooks
  ReadMobScoreArr PT_ROOK, -56, -78, -25, -18, -11, 26, -5, 55, -4, 70, -1, 81, 8, 109, 14, 120, 21, 128, 23, 143, 31, 154, 32, 160, 43, 165, 49, 168, 59, 169
  ' Queens
  ReadMobScoreArr PT_QUEEN, -40, -35, -25, -12, 2, 7, 4, 19, 14, 37, 24, 55, 25, 62, 40, 76, 43, 79, 47, 87, 54, 94, 56, 102, 60, 111, 70, 116, 72, 118, 73, 122, 75, 128, 77, 130, 85, 133, 94, 136, 99, 140, 108, 157, 112, 158, 113, 161, 118, 174, 119, 177, 123, 191, 128, 199

  ' king safety
  KingAttackWeights(PT_PAWN) = 0: KingAttackWeights(PT_KNIGHT) = 77: KingAttackWeights(PT_BISHOP) = 55: KingAttackWeights(PT_ROOK) = 44: KingAttackWeights(PT_QUEEN) = 11

  ' Threats
  ReadScoreArr ThreatByMinor, 0, 0, 0, 31, 39, 42, 57, 44, 68, 112, 62, 120 'Minor on Defended
  ReadScoreArr ThreatByRook, 0, 0, 0, 24, 38, 71, 38, 61, 0, 38, 51, 38 'Major on Defended
  SetScoreVal ThreatByKing, 24, 89
  SetScoreVal ThreatByRank, 13, 0
  SetScoreVal ThreatBySafePawn, 173, 94
  SetScoreVal ThreatByPawnPush, 45, 40
  SetScoreVal HangingBonus, 69, 36
  SetScoreVal OverloadBonus, 13, 6
  SetScoreVal BackwardMalus, 9, 24
  SetScoreVal DoubledMalus, 11, 56
  SetScoreVal IsolatedMalus, 5, 15
  
  'Outpost minors(Pair MG/EG )[0, 1=supported by pawn]
  SetScoreVal Outpost(PT_KNIGHT, 0), 22, 6: SetScoreVal Outpost(PT_KNIGHT, 1), 36, 12
  SetScoreVal Outpost(PT_BISHOP, 0), 9, 2: SetScoreVal Outpost(PT_BISHOP, 1), 15, 5

  SetScoreVal RookOnFile(0), 18, 7: SetScoreVal RookOnFile(0), 44, 20
  SetScoreVal KingProtector, 6, 6
  
  ReadIntArr PushClose(), 0, 0, 100, 80, 60, 40, 20, 10
  ReadIntArr PushAway(), 0, 5, 20, 40, 60, 80, 90, 100
  ReadIntArr PushToEdges(), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 100, 90, 80, 70, 70, 80, 90, 100, 0, 0, 90, 70, 60, 50, 50, 60, 70, 90, 0, 0, 80, 60, 40, 30, 30, 40, 60, 80, 0, 0, 70, 50, 30, 20, 20, 30, 50, 70, 0, 0, 70, 50, 30, 20, 20, 30, 50, 70, 0, 0, 80, 60, 40, 30, 30, 40, 60, 80, 0, 0, 90, 70, 60, 50, 50, 60, 70, 90, 0, 0, 100, 90, 80, 70, 70, 80, 90, 100
 
  
  InitConnectedPawns
  
End Sub


Public Function ReadMobScoreArr(Pt As enumPieceType, ParamArray pSrc())
  ' Read paramter list into array of type TScore ( MG / EG )
  Dim i As Long
  
  For i = 0 To (UBound(pSrc) - 1) \ 2
    MobilityPt(Pt, i).MG = pSrc(2 * i): MobilityPt(Pt, i).EG = pSrc(2 * i + 1)
  Next

End Function

Public Sub AddScore(ScoreTotal As TScore, ScoreAdd As TScore)
  ScoreTotal.MG = ScoreTotal.MG + ScoreAdd.MG: ScoreTotal.EG = ScoreTotal.EG + ScoreAdd.EG
End Sub

Public Sub MinusScore(ScoreTotal As TScore, ScoreMinus As TScore)
  ScoreTotal.MG = ScoreTotal.MG - ScoreMinus.MG: ScoreTotal.EG = ScoreTotal.EG - ScoreMinus.EG
End Sub

Public Function PSQT64(Pt As enumPieceType, ParamArray pSrc())
  ' Read piece square table as paramter list into array
  ' SF tables are symmetric so file A-D is flipped to E-F
  Dim i As Long, sq As Long, x As Long, y As Long, x2 As Long, y2 As Long, MG As Long, EG As Long

  ' Source table is for file A-D, rank 1-8 > Flip for E-F
  For i = 0 To 31
    MG = pSrc(i * 2): EG = pSrc(i * 2 + 1)
    ' for White
    x = i Mod 4: y = i \ 4: sq = 21 + x + y * 10
    Psqt(COL_WHITE, Pt, sq).MG = MG: Psqt(COL_WHITE, Pt, sq).EG = EG
    ' flip to E-F
    x2 = 7 - x: y2 = y: sq = 21 + x2 + y2 * 10
    Psqt(COL_WHITE, Pt, sq).MG = MG: Psqt(COL_WHITE, Pt, sq).EG = EG
    '
    ' Black
    '
    x2 = x: y2 = 7 - y: sq = 21 + x2 + y2 * 10
    Psqt(COL_BLACK, Pt, sq).MG = MG: Psqt(COL_BLACK, Pt, sq).EG = EG
    x2 = 7 - x: y2 = 7 - y: sq = 21 + x2 + y2 * 10
    Psqt(COL_BLACK, Pt, sq).MG = MG: Psqt(COL_BLACK, Pt, sq).EG = EG
  
  Next

End Function

Public Function PieceSquareVal(ByVal Piece As Long, ByVal Square As Long) As Long
  '--- Piece value for a square
  PieceSquareVal = 0
  If Piece < NO_PIECE And Piece > FRAME Then
   PieceSquareVal = Psqt(PieceColor(Piece), PieceType(Piece), Square).EG
  End If

End Function

Public Sub FillPieceSquareVal()
  Dim Piece As Long, Target As Long

  For Piece = 1 To 13
    For Target = SQ_A1 To SQ_H8
      bEndGame = False
      PsqVal(0, Piece, Target) = PieceSquareVal(Piece, Target)
      bEndGame = True
      PsqVal(1, Piece, Target) = PieceSquareVal(Piece, Target)
    Next
  Next

End Sub


Public Function ReadScoreArr(pDest() As TScore, ParamArray pSrc())
  ' Read parameter list into array of type TScore ( MG / EG )
  Dim i As Long

  For i = 0 To (UBound(pSrc) - 1) \ 2
    pDest(i).MG = pSrc(2 * i): pDest(i).EG = pSrc(2 * i + 1)
  Next

End Function


Public Function GetMin(ByVal X1 As Long, ByVal x2 As Long) As Long
  If X1 <= x2 Then GetMin = X1 Else GetMin = x2
End Function

Public Function GetMax(ByVal X1 As Long, ByVal x2 As Long) As Long
  If X1 >= x2 Then GetMax = X1 Else GetMax = x2
End Function

' Stable sort, keeps order if same sort value, highest on top
Public Sub SortMovesStable(ByVal Ply As Long, ByVal iStart As Long, ByVal iEnd As Long)
  Dim i As Long, j As Long, iMin As Long, IMax As Long, TempMove As TMOVE
  iMin = iStart + 1: IMax = iEnd
  i = iMin: j = i + 1

  Do While i <= IMax
    If Moves(Ply, i).OrderValue > Moves(Ply, i - 1).OrderValue Then
      TempMove = Moves(Ply, i): Moves(Ply, i) = Moves(Ply, i - 1): Moves(Ply, i - 1) = TempMove ' Swap
      If i > iMin Then i = i - 1
    Else
      i = j: j = j + 1
    End If
  Loop

 '--- Check correct sort
 ' For i = iStart To iEnd - 1 ' Check sort order
 '  If Moves(Ply, i).OrderValue < Moves(Ply, i + 1).OrderValue Then Stop
 ' Next
End Sub

Public Function GetSEE(Move As TMOVE) As Long
  ' Returns piece score win for AttackColor ( positive for white or black).
  Dim i               As Long, From As Long, MoveTo As Long, Target As Long
  Dim CapturedVal     As Long, PieceMoved As Boolean
  Dim SideToMove      As enumColor, SideNotToMove As enumColor
  Dim NumAttackers(2) As Long, CurrSgn As Long, MinValIndex As Long, Piece As Long, Offset As Long
  '----
  GetSEE = 0
  If PieceType(Move.Piece) = PT_KING Then GetSEE = Abs(PieceScore(Move.Captured)): Exit Function
  If Move.Castle <> NO_CASTLE Then Exit Function
  From = Move.From: MoveTo = Move.Target: PieceMoved = CBool(Board(From) = NO_PIECE)
  If Not PieceMoved Then
    'If PinnedPieceDir(From, MoveTo, PieceColor(PieceMoved)) <> 0 Then GetSEE = -100000: Exit Function
    Piece = Board(From): Board(From) = NO_PIECE ' Remove piece to open sliding xrays
    If Move.EnPassant = ENPASSANT_CAPTURE Then  ' remove captured pawn not on target field
      If Piece = WPAWN Then Board(MoveTo + SQ_DOWN) = NO_PIECE Else Board(MoveTo + SQ_UP) = NO_PIECE
    End If
  Else
    Piece = Board(MoveTo)
  End If
  Cnt = 0 ' Counter for PieceList array of attackers (both sides)
  Erase Blocker  ' Array to manage blocker of sliding pieces: -1: is blocked, >0: is blocking,index of blocked piece, 0:not blocked/blocking

  ' Find attackers
  For i = 0 To 3 ' horizontal+vertical
    Block = 0: Offset = DirectionOffset(i): Target = MoveTo + Offset
    If Board(Target) = BKING Or Board(Target) = WKING Then
      Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target))
    Else

      Do While Board(Target) <> FRAME
        Select Case Board(Target)
          Case BROOK, BQUEEN, WROOK, WQUEEN
            Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target))
            If Block > 0 Then Blocker(Block) = Cnt: Blocker(Cnt) = -1 '- 1. point to blocked piece index; 2. -1 = blocked
            Block = Cnt
          Case NO_PIECE, WEP_PIECE, BEP_PIECE
            '-- Continue
          Case Else
            Exit Do ' other piece
        End Select

        Target = Target + Offset
      Loop

    End If
  Next

  For i = 4 To 7 ' diagonal
    Block = 0: Offset = DirectionOffset(i): Target = MoveTo + Offset

    Select Case Board(Target)
      Case BKING, WKING
        Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target))
        GoTo lblContinue
      Case WPAWN
        If i = 5 Or i = 7 Then Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target)): Block = Cnt: Target = Target + Offset
      Case BPAWN
        If i = 4 Or i = 6 Then Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target)): Block = Cnt: Target = Target + Offset
    End Select

    Do While Board(Target) <> FRAME
      Select Case Board(Target)
        Case BBISHOP, BQUEEN, WBISHOP, WQUEEN
          Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target))
          If Block > 0 Then Blocker(Block) = Cnt: Blocker(Cnt) = -1 '- 1. point to blocked piece index; 2. -1 = blocked
          Block = Cnt
        Case NO_PIECE, WEP_PIECE, BEP_PIECE
          '-- Continue
        Case Else
          Exit Do ' other piece
      End Select
      Target = Target + Offset
    Loop

lblContinue:
  Next

  ' Knights
  If PieceCnt(WKNIGHT) + PieceCnt(BKNIGHT) > 0 Then
    For i = 0 To 7
      Select Case Board(MoveTo + KnightOffsets(i))
        Case WKNIGHT, BKNIGHT: Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(MoveTo + KnightOffsets(i)))
      End Select
    Next
  End If

  '---<<< End of collecting attackers ---
  ' Count Attackers for each color (non blocked only)
  For i = 1 To Cnt
    If PieceList(i) > 0 And Blocker(i) >= 0 Then NumAttackers(COL_WHITE) = NumAttackers(COL_WHITE) + 1 Else NumAttackers(COL_BLACK) = NumAttackers(COL_BLACK) + 1
  Next

  ' Init swap list
  SwapList(0) = Abs(PieceScore(Move.Captured))
  slIndex = 1
  SideToMove = PieceColor(Move.Piece)
  ' Switch side
  SideNotToMove = SideToMove: If SideToMove = COL_WHITE Then SideToMove = COL_BLACK Else SideToMove = COL_WHITE
  ' If the opponent has no attackers we are finished
  If NumAttackers(SideToMove) = 0 Then
    GoTo lblEndSEE
  End If
  If SideToMove = COL_WHITE Then CurrSgn = 1 Else CurrSgn = -1
  '---- CALCULATE SEE ---
  CapturedVal = Abs(PieceScore(Move.Piece))

  Do
    SwapList(slIndex) = -SwapList(slIndex - 1) + CapturedVal
    ' find least valuable attacker (min value)
    CapturedVal = 99999
    MinValIndex = -1

    For i = 1 To Cnt
      If PieceList(i) <> 0 Then If Sgn(PieceList(i)) = CurrSgn Then If Blocker(i) >= 0 Then If Abs(PieceList(i)) < CapturedVal Then CapturedVal = Abs(PieceList(i)): MinValIndex = i
    Next

    If MinValIndex > 0 Then
      If Blocker(MinValIndex) > 0 Then ' unblock other sliding piece?
        Blocker(Blocker(MinValIndex)) = 0
        'Increase attack number
        If PieceList(Blocker(MinValIndex)) > 0 Then NumAttackers(COL_WHITE) = NumAttackers(COL_WHITE) + 1 Else NumAttackers(COL_BLACK) = NumAttackers(COL_BLACK) + 1
      End If
      PieceList(MinValIndex) = 0 ' Remove from list by setting piece value to zero
    End If
    If CapturedVal = 5000 Then ' King
      If NumAttackers(SideNotToMove) = 0 Then slIndex = slIndex + 1
      Exit Do ' King
    End If
    If CapturedVal = 99999 Then Exit Do
    NumAttackers(SideToMove) = NumAttackers(SideToMove) - 1
    CurrSgn = -CurrSgn: SideNotToMove = SideToMove: If SideToMove = COL_WHITE Then SideToMove = COL_BLACK Else SideToMove = COL_WHITE
    slIndex = slIndex + 1
  Loop While NumAttackers(SideToMove) > 0

  '// Having built the swap list, we negamax through it to find the best
  ' // achievable score from the point of view of the side to move.
  slIndex = slIndex - 1

  Do While slIndex > 0
    'SwapList(slIndex - 1) = GetMin(-SwapList(slIndex), SwapList(slIndex - 1))
    If -SwapList(slIndex) < SwapList(slIndex - 1) Then SwapList(slIndex - 1) = -SwapList(slIndex)
    slIndex = slIndex - 1
  Loop

lblEndSEE:
  If Not PieceMoved Then
    Board(From) = Piece
    If Move.EnPassant = ENPASSANT_CAPTURE Then  ' restore captured pawn not on target field
      If Piece = WPAWN Then Board(MoveTo + SQ_DOWN) = BPAWN Else Board(MoveTo + SQ_UP) = WPAWN
    End If
  End If
  GetSEE = SwapList(0)
End Function

Public Function EvalTo100(ByVal Eval As Long) As Long
  If Abs(Eval) < MATE_IN_MAX_PLY Then EvalTo100 = (Eval * 100&) / CLng(ScoreVal(PT_PAWN).EG) Else EvalTo100 = Eval
End Function

''---------------------------------------------------------------------------
''PrintPos() - board position in ASCII table
''---------------------------------------------------------------------------
'Public Function PrintPos() As String
'  Dim a      As Long, b As Long, c As Long
'  Dim sBoard As String
'  sBoard = vbCrLf
'  If True Then ' Not bCompIsWhite Then  'punto di vista del B (engine e' N)
'    sBoard = sBoard & " ------------------" & vbCrLf
'    For a = 1 To 8
'      sBoard = sBoard & (9 - a) & "| "
'
'      For b = 1 To 8
'        c = 100 - (a * 10) + b
'        sBoard = sBoard & Piece2Alpha(Board(c)) & " "
'      Next
'
'      sBoard = sBoard & "| " & vbCrLf
'    Next
'
'  Else
'
'    For a = 1 To 8
'      sBoard = sBoard & a & vbTab
'
'      For b = 1 To 8
'        c = 10 + (a * 10) - b
'        sBoard = sBoard & Piece2Alpha(Board(c)) & " "
'      Next
'
'      sBoard = sBoard & vbCrLf
'    Next
'
'  End If
'  sBoard = sBoard & " ------------------" & vbCrLf
'  sBoard = sBoard & " " & vbTab & " A B C D E F G H" & vbCrLf
'  PrintPos = sBoard
'End Function

Public Sub SetScoreVal(ScoreSet As TScore, ByVal MGScore As Long, ByVal EGSCore As Long)
  ScoreSet.MG = MGScore: ScoreSet.EG = EGSCore
End Sub

Public Sub AddScoreWithFactor(ScoreTotal As TScore, ScoreAdd As TScore, Factor As Long)
  ScoreTotal.MG = ScoreTotal.MG + ScoreAdd.MG * Factor: ScoreTotal.EG = ScoreTotal.EG + ScoreAdd.EG * Factor
End Sub

Public Sub MinusScoreWithFactor(ScoreTotal As TScore, ScoreMinus As TScore, Factor As Long)
  ScoreTotal.MG = ScoreTotal.MG - ScoreMinus.MG * Factor: ScoreTotal.EG = ScoreTotal.EG - ScoreMinus.EG * Factor
End Sub

Public Function InitConnectedPawns()
  ' SF6
  Dim Seed(8) As Long, Opposed As Long, Phalanx As Long, Support As Long, r As Long, V As Long, x As Long
  ReadLngArr Seed(), 0, 0, 13, 24, 18, 76, 100, 175, 330

  For Opposed = 0 To 1
    For Phalanx = 0 To 1
      For Support = 0 To 2
        For r = 2 To 7
          If Phalanx > 0 Then x = (Seed(r + 1) - Seed(r)) / 2 Else x = 0
          V = 17 * Support
          V = V + Seed(r)
          If Phalanx > 0 Then V = V + (Seed(r + 1) - Seed(r)) \ 2
          If Opposed > 0 Then V = V / 2 ' >>  operator for opposed in VB: /2
          Connected(Opposed, Phalanx, Support, r).MG = V
          Connected(Opposed, Phalanx, Support, r).EG = V * ((r - 1) - 2) \ 4 ' rank r ist zero based in C, so (r-1)
        Next
      Next
    Next
  Next
End Function

Public Function AdjacentFilesLL(FileNum As Long) As LongLong
 ' returns bitboard for fiels left and right of FileNum (1-8 = A-H)
 Select Case FileNum
 Case 1: AdjacentFilesLL = FileB_LL
 Case 2: AdjacentFilesLL = FileA_LL Or FileC_LL
 Case 3: AdjacentFilesLL = FileB_LL Or FileD_LL
 Case 4: AdjacentFilesLL = FileC_LL Or FileE_LL
 Case 5: AdjacentFilesLL = FileD_LL Or FileF_LL
 Case 6: AdjacentFilesLL = FileE_LL Or FILEG_LL
 Case 7: AdjacentFilesLL = FileF_LL Or FileH_LL
 Case 8: AdjacentFilesLL = FILEG_LL
 End Select
End Function


Public Function OpenFiles() As Long
 ' Return number of files without pawns
 Dim InBoardLL As LongLong
 InBoardLL = PiecesLL(COL_WHITE, PT_PAWN) Or PiecesLL(COL_BLACK, PT_PAWN)
 OpenFiles = Abs((InBoardLL And FileA_LL) = 0) + Abs((InBoardLL And FileB_LL) = 0) + Abs((InBoardLL And FileC_LL) = 0) + Abs((InBoardLL And FileD_LL) = 0) + _
             Abs((InBoardLL And FileE_LL) = 0) + Abs((InBoardLL And FileF_LL) = 0) + Abs((InBoardLL And FILEG_LL) = 0) + Abs((InBoardLL And FileH_LL) = 0)
End Function

Public Function SideToMove() As enumColor
  If bWhiteToMove Then SideToMove = COL_WHITE Else SideToMove = COL_BLACK
End Function

Public Function Endgame_KXK(StrongCol As enumColor, NonPawnMaterialStrongSide As Long) As Long
  Dim WinnerKingSq As Long, LoserKingSq As Long, WeakCol As enumColor, Result As Long
  If StrongCol = COL_WHITE Then
    WinnerKingSq = WKingLoc: LoserKingSq = BKingLoc: WeakCol = COL_BLACK
  Else
    WinnerKingSq = BKingLoc: LoserKingSq = WKingLoc: WeakCol = COL_WHITE
  End If
  
  'If Board(SQ_F4) = WKING And Board(SQ_H2) = BKING Then Stop
  Result = PopCnt64LL(PieceMovesLL(WKING, SqToBit(LoserKingSq)) And Not AttacksForColLL(StrongCol)) - Abs(InCheck())
  If Result = 1 Then
    Result = 800
  ElseIf Result = 2 Then Result = 500
  ElseIf Result = 3 Then Result = 300
  ElseIf Result = 4 Then Result = 200
  ElseIf Result = 5 Then Result = 100
  End If
  
  Result = Result + NonPawnMaterialStrongSide _
        + PieceCntCol(StrongCol, PT_PAWN) * PieceTypeValue(PT_PAWN) _
        + 3 * PushToEdges(LoserKingSq) _
        + PushClose(MaxDistance(WinnerKingSq, LoserKingSq))
  If InCheck() Then Result = Result + 250
  If StrongCol = SideToMove Then Endgame_KXK = Result Else Endgame_KXK = -Result
End Function

#End If


