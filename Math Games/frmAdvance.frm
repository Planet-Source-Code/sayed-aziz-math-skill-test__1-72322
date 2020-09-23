VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Math Logic Game"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAdvance.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5985
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAnswer1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtAnswer 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4080
      Width           =   5055
   End
   Begin VB.TextBox txtAns 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   12
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Math Logic Game"
      Top             =   120
      Width           =   4935
   End
   Begin VB.TextBox txtBal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtCurr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Left            =   840
      Top             =   2160
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7320
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   2160
   End
   Begin VB.CommandButton cmdQuit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Picture         =   "frmAdvance.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Picture         =   "frmAdvance.frx":3D80
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080C0FF&
      X1              =   5400
      X2              =   8400
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Image Image1 
      Height          =   2010
      Left            =   5640
      Picture         =   "frmAdvance.frx":463B
      Top             =   840
      Width           =   2670
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Points In Stock "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   4800
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Current Points"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   5160
      Width           =   3615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Total Points Upto This Level "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   5520
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080C0FF&
      Height          =   5415
      Left            =   5400
      Top             =   480
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      Height          =   1170
      Left            =   120
      Top             =   4680
      Width           =   5055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Left"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varAge1 As Integer, varAge2 As Integer, varAge3 As Integer, X As Integer, Y As Integer
Dim varrnd2, Q As Integer, varQuit As Boolean
Private Sub ScoreBoard()

 On Error Resume Next
  
  sndPlaySound vbNullString, sndAsync
  
  ConnectAccessDb
  
  Set rs = dbs.OpenRecordset("SELECT * FROM tblScore", dbOpenDynaset)
  
  With rs
    
   If usrId > 0 Then GoTo 10
   
    If rs.BOF Then
     usrId = 1
    
    Else
     rs.MoveLast
     usrId = rs!kidid + 1
     
    End If
    
10    .FindFirst "Kidid=" & usrId
    
    If .NoMatch = True Then
        
        .AddNew
        !kidid = usrId
        !kidname = KidsName
        !kidscore = varPoints
        !kdate = Now
        .Update
    
    Else
        
        .Edit
        !kidscore = varPoints
        !kdate = Now
        .Update
        
    End If
    
  End With
  
   CloseAccessDb

   If varQuit = True Then
      Load Form5
      Form5.Show
   Else
      End
   End If
   
End Sub
Private Sub cmdQuit_Click()
    Call Form_KeyPress(27)
End Sub
Private Sub Command1_GotFocus()
 
 Me.txtAns.Enabled = False
 
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
 
  If KeyAscii = 27 Then
        Dim Response
        
        Me.Timer1.Enabled = False
        
        Response = MsgBox("Do You Wish To Quit Game ?", vbExclamation + vbYesNo, "Quit !!")
         
        If Response = vbYes Then
           
           If varPoints > 0 Then
            
            Unload Me
            ScoreBoard
            varQuit = True
            
           Else
            
            sndPlaySound vbNullString, sndAsync
            Unload Me
            End
            
           End If
            
        Else
           If Me.Text2 > 0 Then Me.Timer1.Enabled = True
        End If
    
    End If
    
End Sub
Private Sub Command1_Click()
  
 Dim strSum As String, varTest As Integer, varTest1 As Integer, varTest2 As Integer, varTest3 As Integer
 sndPlaySound vbNullString, sndAsync
 
 Me.Caption = "Advance Level Stage " & Q & " / 5 "
 Me.Timer2.Enabled = False
 
 Image1.Picture = LoadPicture(App.Path & "\Child.jpg")
 
 Me.txtAnswer = ""
 
 Me.txtAns.Visible = True
 Me.txtAns.Enabled = True
 Me.txtAns.SetFocus
 Me.txtAns = Empty
 Me.txtAnswer1.Visible = True
 Me.txtAnswer1 = Empty
 
 Me.Command1.Enabled = False
  
 Randomize (varrnd2)
 varrnd1 = Int((45 * Rnd) + 1)
 varrnd2 = varrnd1
 
  ConnectAccessDb
  Set rs = dbs.OpenRecordset("SELECT * FROM tblLogic", dbOpenDynaset)
  
  With rs
    
    .MoveFirst
    .FindFirst "sumid=" & varrnd1
    
  Select Case varrnd1
    
    Case 1 To 45
        
        varAge3 = rs!ans
        
        Me.Text1 = rs!Sum
        
        Me.txtAnswer1 = IIf(IsNull(rs!ans1), Empty, rs!ans1)
        
   End Select
  
  End With
  
   CloseAccessDb
   
   Me.txtBal = varPoints
   Me.txtCurr = 0
   Me.txtTot = varPoints + Val(Me.txtCurr)
     
   Me.Text1.Visible = True
   Me.Timer1.Enabled = True
   Me.Timer1.Interval = 1000
   
   Me.Text2 = 25
     
End Sub
Private Sub Form_Load()
 
 sndPlaySound vbNullString, sndAsync
 
 Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
 
 Me.Text1.Visible = False
 Me.txtAns.Visible = False
 Me.txtAnswer1.Visible = False
 
 varQuit = False
 Q = 1
 
End Sub
Private Sub Form_Unload(Cancel As Integer)
  sndPlaySound vbNullString, sndAsync
End Sub
Private Sub Text1_GotFocus()
 If Me.txtAns.Enabled = True Then Me.txtAns.SetFocus
End Sub
Private Sub Timer1_Timer()
    
    Me.Text2 = Text2 - 1
 
 If Me.Text2 = 0 Then
    
    Me.Timer1.Enabled = False
              
    Me.Timer2.Enabled = True
    Me.Timer2.Interval = 150
            
    varChk = False
    varPoints = varPoints - 1500
    Me.txtCurr = -1500
    Me.txtTot = Val(Me.txtBal) + Val(Me.txtCurr)
    varPoints = Me.txtTot
    NoisePlay WavMov, SND_ASYNC Or SND_LOOP
    Me.Command1.Enabled = True
    Me.Command1.SetFocus
    Me.txtAns.Enabled = False
    
 End If
    
End Sub
Private Sub Timer2_Timer()
  
  If varChk = True Then
    
    Image1.Picture = LoadPicture(App.Path & "\animation\P0" & X & ".bmp")
    X = X + 1
    
    If X = 16 Then
        X = 0
    End If
  
  Else
   
    Image1.Picture = LoadPicture(App.Path & "\animation\L0" & Y & ".bmp")
    Y = Y + 1
    
    If Y = 3 Then
        Y = 0
    End If
    
  End If
  
End Sub
Private Sub txtAns_Change()
 If Not IsNumeric(Me.txtAns) Then Me.txtAns = Empty
End Sub
Private Sub txtAns_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Len(Trim(Me.txtAns)) > 0 Then
       
       Me.Command1.Enabled = True
       Me.Command1.SetFocus
       Me.Timer1.Enabled = False
       
       If Val(Me.txtAns) = varAge3 Then
        
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 150
        varChk = True
    
        NoisePlay WavClaps, SND_ASYNC Or SND_LOOP
        Q = Q + 1
    
        Me.txtBal = Val(varPoints)
        varPoints = (Q * 1500) + (150 * Me.Text2)
        Me.Label6.Caption = "Current Points " & Q * 1500 & " + " & Me.Text2 & " (i.e. Time Left) X 150"
             
        Me.txtCurr = varPoints
        Me.txtTot = Val(Me.txtBal) + Val(Me.txtCurr)
        varPoints = Me.txtTot
        
        Me.Command1.Enabled = True
        Me.Command1.SetFocus
        
        If Q > 5 Then
            
            varQuit = True
            sndPlaySound vbNullString, sndAsync
            Unload Me
            Call ScoreBoard
        
        End If
    
       Else
  
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 150
        varChk = False
        Me.txtAnswer = "Oh No .... Wrong Answer .... Correct Answer is " & varAge3
    
        varPoints = varPoints - 1500
        Me.txtCurr = -1500
        Me.txtTot = Val(Me.txtBal) + Val(Me.txtCurr)
        varPoints = Me.txtTot
        NoisePlay WavMov, SND_ASYNC Or SND_LOOP
        
       End If
    
    End If
    
End Sub
