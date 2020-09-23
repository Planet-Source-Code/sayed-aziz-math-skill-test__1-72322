VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1215
   ClientLeft      =   3990
   ClientTop       =   2205
   ClientWidth     =   2985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1215
      Left            =   0
      Picture         =   "frmLevel.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   2910
      TabIndex        =   0
      Top             =   0
      Width           =   2970
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   840
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varPic As Integer
Private Sub Form_Load()
 
 Me.Timer1.Enabled = True
 Me.Timer1.Interval = 1000
 varPic = 1
 
End Sub
Private Sub Timer1_Timer()
  
  If Me.Visible = True Then Me.SetFocus
  
  If varPic = 1 Then
     
     Me.Picture1.Picture = LoadPicture(App.Path & "\Animation\SelfAnim1.bmp")
     Me.Width = Me.Width - 150
     Me.Height = Me.Height - 130
     varPic = varPic + 1
     
  ElseIf varPic = 2 Then
     
     Me.Picture1.Picture = LoadPicture(App.Path & "\Animation\SelfAnim2.bmp")
     Me.Width = Me.Width - 440
     Me.Height = Me.Height - 150
     varPic = varPic + 1
     
  ElseIf varPic = 3 Then
     
     Me.Picture1.Picture = LoadPicture(App.Path & "\Animation\SelfAnim3.bmp")
     Me.Width = Me.Width - 575
     Me.Height = Me.Height - 250
     varPic = varPic + 1
     
     
  ElseIf varPic = 4 Then
    
    Me.Picture1.Picture = LoadPicture(App.Path & "\Animation\SelfAnim4.bmp")
    Me.Width = Me.Width - 600
    Me.Height = Me.Height - 175
    varPic = varPic + 1
    
    
  ElseIf varPic > 4 Then
  
     Me.Picture1.Picture = LoadPicture(App.Path & "\Animation\SelfAnim5.bmp")
     Me.Width = Me.Width - 600
     Me.Height = Me.Height - 250
     Me.Timer1.Enabled = False
     Me.Timer1.Interval = 0
     
     Unload Me
     Form1.SetFocus
     
  End If
  
End Sub
