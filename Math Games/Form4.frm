VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                        Math Skill"
   ClientHeight    =   3990
   ClientLeft      =   1050
   ClientTop       =   1935
   ClientWidth     =   5955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5955
   Begin VB.CommandButton cmdReset 
      Height          =   435
      Left            =   4560
      Picture         =   "Form4.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2880
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Math Skill Test"
      ClipControls    =   0   'False
      Height          =   975
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "¸"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   15.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   320
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   200
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   435
      Left            =   3600
      Picture         =   "Form4.frx":29D0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   3780
      Left            =   120
      Picture         =   "Form4.frx":39EA
      ScaleHeight     =   3720
      ScaleWidth      =   3135
      TabIndex        =   14
      Top             =   120
      Width           =   3200
   End
   Begin VB.CommandButton Command1 
      Height          =   600
      Left            =   3720
      Picture         =   "Form4.frx":291CC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1715
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ver 1.0"
      Height          =   200
      Left            =   3840
      TabIndex        =   13
      Top             =   3520
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      DrawMode        =   16  'Merge Pen
      Height          =   3750
      Left            =   3360
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Your Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   225
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReset_Click()
  
 Dim Response
 
 ConnectAccessDb
 Set rs = dbs.OpenRecordset("SELECT * FROM tblScore", dbOpenDynaset)
        
 If rs.BOF = True Then
     
        MsgBox "Nothing To Reset", vbInformation, "Blank Scoreboard !!"
        GoTo 10
 
 Else
     
     Response = MsgBox("Do you want to Delete All Scores ?", vbQuestion + vbYesNo, "Confirm !!")
 
    If Response = vbYes Then
    
        rs.MoveFirst
    
        Do While Not rs.EOF
 
            rs.Delete
            rs.MoveNext
 
        Loop
 
    End If
 
 End If
 
10  CloseAccessDb
    Me.Text1.SetFocus
    
End Sub
Private Sub Command1_Click()

  WavClaps = App.Path & "\claps.wav"
  WavClaps1 = App.Path & "\claps1.wav"
  WavMov = App.Path & "\Move.wav"
  
  If IsNull(Me.Text1) Or Me.Text1 = "" Then
     
        Me.Text1.SetFocus
  
  Else
         
      KidsName = StrConv(Me.Text1.Text, vbProperCase)
      Me.Text1 = Empty
            
     If Me.Option1 = True Then
        
        Z = 1
        Unload Me
        Load Form1
        Form1.Show
        
     ElseIf Me.Option2 = True Then
        
        Z = 4
        Unload Me
        Load Form1
        Form1.Show
        
     ElseIf Me.Option3 = True Then
        
        Z = 7
       
        Unload Me
        Load Form2
        Form2.Show
        
     ElseIf Me.Option4 = True Then
        
        Z = 10
       
        Unload Me
        Load Form2
        Form2.Show
        
     End If
     
  End If
  
End Sub
Private Sub Command2_Click()
 
 Unload Me
 End
 
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then Call Command2_Click
    
End Sub
Private Sub Form_Load()
 
 Dim i As Integer, i1 As Integer
 
 sndPlaySound vbNullString, sndAsync
 
 Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
 
End Sub
Private Sub Frame1_Click()
 
 If IsNull(Me.Text1) Or Len(Trim(Me.Text1)) = 0 Then
     Me.Text1.SetFocus
 End If
 
End Sub
Private Sub Option1_Click()
 
 If IsNull(Me.Text1) Or Len(Trim(Me.Text1)) = 0 Then
     Me.Text1.SetFocus
 End If
 
End Sub
Private Sub Option1_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 And Me.Option1.Value = True Then
    Call Command1_Click
    'Me.Command1.SetFocus
 End If
 
End Sub

Private Sub Option2_Click()
 
 If IsNull(Me.Text1) Or Len(Trim(Me.Text1)) = 0 Then
     
     Me.Text1.SetFocus
     Me.Option1.Value = True
     Me.Option2.Value = False
     
 End If
 
End Sub
Private Sub Option2_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 And Me.Option2.Value = True Then
    Call Command1_Click
    'Me.Command1.SetFocus
 End If

End Sub
Private Sub Option3_Click()
 
 If IsNull(Me.Text1) Or Len(Trim(Me.Text1)) = 0 Then
     
     Me.Text1.SetFocus
     Me.Option1.Value = True
     Me.Option3.Value = False
     
 End If
 
End Sub
Private Sub Option3_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 And Me.Option3.Value = True Then
    Call Command1_Click
    'Me.Command1.SetFocus
    
 End If

End Sub
Private Sub Option4_Click()
 
 If IsNull(Me.Text1) Or Len(Trim(Me.Text1)) = 0 Then
     
     Me.Text1.SetFocus
     Me.Option1.Value = True
     Me.Option4.Value = False
     
 End If

End Sub
Private Sub Option4_KeyPress(KeyAscii As Integer)
  
 If KeyAscii = 13 And Me.Option4.Value = True Then
    Call Command1_Click
    'Me.Command1.SetFocus
 End If

End Sub
Private Sub Picture1_Click()
 
 Me.Text1.SetFocus
 
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Len(Trim(Me.Text1)) > 0 Then
        Me.Option1.SetFocus
    End If
    
End Sub
Private Sub Text1_LostFocus()
 
 Me.Text1 = StrConv(Me.Text1, vbProperCase)
 
End Sub
