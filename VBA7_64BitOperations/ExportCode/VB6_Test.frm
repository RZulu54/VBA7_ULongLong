VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB6 emulate 64 bit usigned integer bit operations"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   13275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShiftRight2 
      Caption         =   "ShiftRight1 pattern B"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdShiftRight1 
      Caption         =   "ShiftRight1 pattern A"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cdmShiftLeft2 
      Caption         =   "ShiftLeft1 pattern B"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdShiftLeft 
      Caption         =   "ShiftLeft1 pattern A"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.ListBox lstOut 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7035
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   11895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdShiftLeft_Click()
  Test_ShiftLeft1
End Sub

Private Sub cdmShiftLeft2_Click()
 Test_ShiftLeft2
End Sub

Private Sub cmdShiftRight1_Click()
  Test_ShiftRight1
End Sub

Private Sub cmdShiftRight2_Click()
  Test_ShiftRight2
End Sub


Public Sub Test_ShiftLeft1()
  Test32VB6 "SHIFTLEFT1"
End Sub

Public Sub Test_ShiftLeft2()
  Test32VB6 "SHIFTLEFT2"
End Sub

Public Sub Test_ShiftRight1()
  Test32VB6 "SHIFTRIGHT1"
End Sub

Public Sub Test_ShiftRight2()
  Test32VB6 "SHIFTRIGHT2"
End Sub

Public Sub Test32VB6(ByVal sTestCase As String)
  Dim i As Long, j As Long
  Dim Bit64L As TBit64, Bit64aL As TBit64, BitShifts As Long
  
  lstOut.Clear ' clear the test output box
  
  'Debug.Print String(255, vbNewLine) ' Clear Debug
  
  BitShifts = 1
  
  Init32Bit
  
  Select Case sTestCase
  Case "SHIFTLEFT1", "SHIFTLEFT2", "SHIFTRIGHT1", "SHIFTRIGHT2"
    '--- Test ShiftLeft , ShiftRight> fill start pattern
    If Right$(sTestCase, 1) = "1" Then j = 0 Else j = 1 ' start with bit 0 or bit 1, end with bit 62 or bit 32
    Clear64 Bit64L
    For i = 0 To 63
      If i Mod 2 = j Then
        SetBit64 Bit64L, i
      End If
    Next
    lstOut.AddItem "0_______8_______16______24______32______40______48______56_____63 (Bit Count  LeftPos RightPos)"
    ShowBit64L Bit64L
    
    Bit64aL = Bit64L
    For i = 0 To 63
      If Left$(sTestCase, 9) = "SHIFTLEFT" Then
        Bit64aL = BitsShiftLeft64(Bit64aL, BitShifts): ShowBit64L Bit64aL
      '  Bit64aL = ShiftLeft64(Bit64aL): ShowBit64L Bit64aL
      ElseIf Left$(sTestCase, 10) = "SHIFTRIGHT" Then
        Bit64aL = BitsShiftRight64(Bit64aL, BitShifts): ShowBit64L Bit64aL
       ' Bit64aL = ShiftRight64(Bit64aL): ShowBit64L Bit64aL
      End If
    Next i
    'Debug.Print String(90, "=")
  
 
  End Select

End Sub

Private Sub ShowBit64L(bbInL As TBit64)
 Dim sLine As String
 sLine = GetBitList64L(bbInL) & "  (PopCnt:" & Right$("  " & PopCnt64(bbInL), 2) & "  LSB:" & Right$("  " & Lsb64(bbInL), 2) & "  RSB:" & Right$("  " & Rsb64(bbInL), 2) & ")"
 lstOut.AddItem sLine
End Sub

