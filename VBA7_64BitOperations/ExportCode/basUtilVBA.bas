Attribute VB_Name = "basUtilVBA"
'=========================================================================
'= UtilVBAbas:
'= functions for VBA GUI ( VBA= Visual Basic for Application in MS-Office)
'==========================================================================
Option Explicit

Public Const TEST_MODE As Boolean = True
Public ThisApp As Object ' Office object: Excel, Word,...
Public psGameFile As String
Public LastInfoNodes As Long

Public psLastFieldClick As String
Public psLastFieldMouseDown As String
Public psLastFieldMouseUp As String

Public SetupBoardMode As Boolean  ' manual board setup using GUI
Public SetupPiece As Long

' GUI colors
Public WhiteSqCol As Long
Public BlackSqCol As Long
Public BoardFrameCol As Long

Public plFieldFrom As Long, plFieldTarget As Long
Public psFieldFrom As String, psFieldTarget As String
Dim plFieldFromColor As Long, plFieldTargetColor As Long
Dim psMove As String


Sub run_ChessBrainX()
 ' Main
 
 #If VBA7 And Win64 Then 'Note: Win64 = Office64 bit (not Windows 64 bit)
  InitEngine
  
  Init64Bit ' general 64 bit functions for LongLong datatype
  InitFileRankLL
  
  ' 64 bit chess bitboards
  InitAttackHelperBoards
  InitAttacksFromSquareLL
  
  InitGame
  
  frmChessX.Show
 #Else
   MsgBox "for 64bit Office only"
 #End If
End Sub


#If VBA7 And Win64 Then 'Note: Win64 = Office64 bit (not Windows 64 bit)

Public Sub SetVBAPathes()
 
End Sub

Public Sub DoFieldClicked()
  ' square click handling: 1. click: select FROM square, 2. click: select TARGET square => do move
  Dim bIsLegal As Boolean, NumLegalMoves As Long, FieldPos As Long, FieldTarget As Long
  Dim sPromotePiece As String, lResult As Long
  
  '--- Setup board mode:  if square not empty: 1 click: white piece, 2. click: black piece, 3. click: clear field
  If SetupBoardMode Then
    If Trim(psLastFieldClick) <> "" Then
      If SetupPiece > 0 Then
        psFieldFrom = psLastFieldClick
        plFieldFrom = Val("0" & Mid(psLastFieldClick, Len("Square") + 1))
        FieldPos = FieldNumToBoardPos(plFieldFrom)
        
        If Board(FieldPos) = NO_PIECE Or (PieceType(Board(FieldPos)) <> PieceType(SetupPiece)) Then
          Board(FieldPos) = SetupPiece
        ElseIf PieceColor(Board(FieldPos)) = COL_WHITE Then
          If PieceColor(SetupPiece) = COL_WHITE Then
             Board(FieldPos) = SetupPiece + 1 ' Black piece, same type
          Else
             Board(FieldPos) = NO_PIECE
          End If
        ElseIf PieceColor(Board(FieldPos)) = COL_BLACK Then
          If PieceColor(SetupPiece) = COL_BLACK Then
             Board(FieldPos) = SetupPiece - 1 ' white piece, same type
          Else
             Board(FieldPos) = NO_PIECE
          End If
        Else
          ' Clear
           Board(FieldPos) = NO_PIECE
        End If
        frmChessX.ShowBoard
        DoEvents
      End If
    End If
    Exit Sub
  End If
  
  ' Move input
  If Trim(psLastFieldClick) <> "" Then
    If plFieldFrom = 0 Then
    
      '--- First click: Field from
      psFieldFrom = psLastFieldClick
      plFieldFrom = Val("0" & Mid(psLastFieldClick, Len("Square") + 1))
      FieldPos = FieldNumToBoardPos(plFieldFrom)
      If Board(FieldPos) < NO_PIECE Then
        '-- check color to move
        If bWhiteToMove And Board(FieldPos) Mod 2 <> 1 Or _
          Not bWhiteToMove And Board(FieldPos) Mod 2 <> 0 Then
          '--- wrong color
          SendCommand "Wrong color! "
          plFieldFrom = 0
          ResetGUIFieldColors
        Else
          frmChessX.Controls(psLastFieldClick).BackColor = &HFF8080
          ResetGUIFieldColors
          ShowLegalMovesForPiece FieldNumToCoord(plFieldFrom)
        End If
      Else
        ' ignore empty field
        plFieldFrom = 0
        ResetGUIFieldColors
      End If
      
    Else
      
      '--- Second click: Field target
      If psLastFieldClick = psFieldFrom Then
         ResetGUIFieldColors
         DoEvents
         plFieldFrom = 0
      Else
        psFieldTarget = psLastFieldClick
        plFieldTarget = Val("0" & Mid(psLastFieldClick, Len("Square") + 1))
        frmChessX.Controls(psLastFieldClick).BackColor = &HC0FFC0
        DoEvents
        Sleep 250
        '--- Check player move
        bIsLegal = CheckGUIMoveIsLegal(FieldNumToCoord(plFieldFrom), FieldNumToCoord(plFieldTarget), NumLegalMoves)
        If bIsLegal Then
          ' Promotion?
          sPromotePiece = "": FieldPos = FieldNumToBoardPos(plFieldFrom): FieldTarget = FieldNumToBoardPos(plFieldTarget)
          If (Board(FieldPos) = WPAWN And Rank(FieldTarget) = 8) Or (Board(FieldPos) = BPAWN And Rank(FieldTarget) = 1) Then
            lResult = MsgBox(Translate("Promote to queen?"), vbYesNo) ' or Knight
            If lResult = vbYes Then sPromotePiece = "q" Else sPromotePiece = "n"
          End If
          '--- Send move to Engine
          psMove = FieldNumToCoord(plFieldFrom) & FieldNumToCoord(plFieldTarget) & sPromotePiece & vbLf
          ParseCommand psMove
          frmChessX.txtEvalMoves.Text = ""
          frmChessX.ShowMoveList
          frmChessX.ShowBoard
        Else
          If NumLegalMoves = 0 Then
            If InCheck() Then
              MsgBox "Mate!"
            Else
              MsgBox "No legal move -> Draw!!!"
            End If
          Else
            SendCommand "Illegal move: " & FieldNumToCoord(plFieldFrom) & FieldNumToCoord(plFieldTarget) & " !!!"
          End If
        End If
        
        'Reset
        plFieldFrom = 0: plFieldTarget = 0
        ResetGUIFieldColors
        
        If bIsLegal And frmChessX.chkAutoThink = True Then
          DoEvents
          frmChessX.cmdThink_Click
          DoEvents
        End If
      End If
    End If
  Else
   ResetGUIFieldColors
  End If
  DoEvents
End Sub


Public Function FieldNumToBoardPos(ByVal ilFieldNum As Long) As Long
   Dim s As String
   s = FieldNumToCoord(ilFieldNum)
   FieldNumToBoardPos = FileRev(Left(s, 1)) + RankRev(Mid(s, 2, 1))
End Function


Public Function CheckGUIMoveIsLegal(MoveFromText, MoveTargetText, oLegalMoves As Long) As Boolean
  ' Input: "e2", "e4", Output:  oLegalMoves:Number of Legal Moves
  Dim a As Long, NumMoves As Long, From As Long, Target As Long
  CheckGUIMoveIsLegal = False
  
  Ply = 1
  oLegalMoves = GenerateLegalMoves(NumMoves)
  If oLegalMoves > 0 Then
    From = FileRev(Left(MoveFromText, 1)) + RankRev(Mid(MoveFromText, 2, 1))
    Target = FileRev(Left(MoveTargetText, 1)) + RankRev(Mid(MoveTargetText, 2, 1))
    
    For a = 0 To NumMoves - 1
       If Moves(1, a).From = From And Moves(1, a).Target = Target Then
          CheckGUIMoveIsLegal = Moves(1, a).IsLegal: Exit For
       End If
    Next a
  End If
End Function

Public Sub ShowLegalMovesForPiece(MoveFromText)
  ' Input: square as text "e2"
  Dim a As Long, NumMoves As Long, From As Long, Target As Long
  Dim NumLegalMoves As Long, ctrl As Control, bFound As Boolean
  
  Ply = 1: bFound = False
  NumLegalMoves = GenerateLegalMoves(NumMoves)
  From = FileRev(Left(MoveFromText, 1)) + RankRev(Mid(MoveFromText, 2, 1))
  If NumLegalMoves = 0 Then
    SendCommand "No legal moves!"
  Else
    For Each ctrl In frmChessX.Controls
      Target = Val("0" & ctrl.Tag)
      If Target > 0 Then
        For a = 0 To NumMoves - 1
         If Moves(1, a).From = From And Moves(1, a).Target = Target And Moves(1, a).IsLegal Then
           ctrl.BackColor = &HC0FFC0
           bFound = True
         End If
        Next a
      End If
    Next ctrl
    If Not bFound Then
      SendCommand "No legal move for this piece!"
    End If
  End If

End Sub

Public Sub ResetGUIFieldColors()
 Dim x As Long, y As Long, bBackColorIsWhite As Boolean, i As Long
 
 bBackColorIsWhite = False
 
 For y = 1 To 8
  For x = 1 To 8
    i = x + (y - 1) * 8
    With frmChessX.fraBoard.Controls("Square" & i)
      If bBackColorIsWhite Then
       If .BackColor <> WhiteSqCol Then .BackColor = WhiteSqCol
      Else
       If .BackColor <> BlackSqCol Then .BackColor = BlackSqCol
      End If
    End With
    bBackColorIsWhite = Not bBackColorIsWhite
  Next x
  bBackColorIsWhite = Not bBackColorIsWhite
 Next y
End Sub

Public Sub ShowBitBoard64(ByVal InBoardLL As LongLong)
  ' Input: 64 bit mask > all bits set are highlighted on board form
  Dim ctrl As Control, Square As Long

  ResetGUIFieldColors

  For Each ctrl In frmChessX.Controls
    Square = Val("0" & ctrl.Tag)
    If Square > 0 Then
      If InBoardLL And Bit64ValueLL(SqToBit(Square)) Then
       ctrl.BackColor = &HC0FFC0
      End If
    End If
  Next
End Sub
#End If

