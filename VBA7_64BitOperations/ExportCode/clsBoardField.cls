VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBoardField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--- Catch Events for board square image controls (VBA has no support for control arrays like VB6)
Public WithEvents ImageEvents As MSForms.Image
Attribute ImageEvents.VB_VarHelpID = -1
Public Name As String

#If VBA7 And Win64 Then 'Note: Win64 = Office64 bit (not Windows 64 bit)

Public Sub SetBoardField(ctl As MSForms.Image)
    Set ImageEvents = ctl
End Sub
 
Private Sub ImageEvents_Click()
  psLastFieldClick = Me.Name
  DoFieldClicked
  DoEvents
End Sub


Private Sub ImageEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  psLastFieldMouseDown = Me.Name
End Sub

#End If
