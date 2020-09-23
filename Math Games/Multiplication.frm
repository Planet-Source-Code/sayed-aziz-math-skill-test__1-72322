VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   "
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8265
   FillColor       =   &H00E0E0E0&
   Icon            =   "Multiplication.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Height          =   450
      Left            =   6360
      Picture         =   "Multiplication.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   850
   End
   Begin VB.TextBox txtResult 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtSecond 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtFirst 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Height          =   495
      Left            =   6000
      Picture         =   "Multiplication.frx":2CE4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   720
      Picture         =   "Multiplication.frx":359F
      ScaleHeight     =   1095
      ScaleWidth      =   975
      TabIndex        =   10
      Top             =   1320
      Width           =   975
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
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4560
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
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4920
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
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   120
      Top             =   3720
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7200
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   3720
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   6240
      Picture         =   "Multiplication.frx":6B7D
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080C0FF&
      X1              =   5160
      X2              =   8160
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      X1              =   1920
      X2              =   3720
      Y1              =   3800
      Y2              =   3800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      X1              =   1920
      X2              =   3720
      Y1              =   3700
      Y2              =   3700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      X1              =   1920
      X2              =   3720
      Y1              =   2680
      Y2              =   2680
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
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      Height          =   1170
      Left            =   45
      Top             =   4485
      Width           =   5055
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
      Left            =   120
      TabIndex        =   5
      Top             =   4920
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
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080C0FF&
      Height          =   5415
      Left            =   5160
      Top             =   240
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   2010
      Left            =   5400
      Picture         =   "Multiplication.frx":A8E3
      Top             =   1080
      Width           =   2670
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Time Left :"
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
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   270
      Left            =   3960
      TabIndex        =   0
      Top             =   4020
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      DrawMode        =   5  'Not Copy Pen
      FillStyle       =   0  'Solid
      Height          =   3765
      Left            =   75
      Top             =   240
      Width           =   5000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X1 As Integer, Y1 As Integer, Z1 As Integer
Private Sub cmdPlay_Click()
    
    Dim retVal
    Me.Timer2.Enabled = False
    Me.Image2.Visible = False
    
    sndPlaySound vbNullString, sndAsync
    DoEvents
    
    Me.Image1.Picture = LoadPicture(App.Path & "\Child" & ".jpg")
    
    Me.txtResult.Enabled = True
    Me.txtResult = Empty
 
    If Z >= 7 And Z <= 9 Then
     
     Me.Tag = "Multiplication"
     Me.txtBal = varPoints
     Me.txtCurr = 0
     Me.txtTot = varPoints + Val(Me.txtCurr)
     Me.Picture1.Picture = LoadPicture(App.Path & "\Multiply" & ".bmp")
    
     Call fncMultiply
  
    ElseIf Z >= 10 And Z <= 12 Then
     
     Me.Tag = "Division"
     Me.txtBal = varPoints
     Me.txtCurr = 0
     Me.txtTot = varPoints + Val(Me.txtCurr)
     Me.Picture1.Picture = LoadPicture(App.Path & "\Div" & ".bmp")
     Call fncDivision
     
    Else
     
     Unload Me
     
     Load Form3
     Form3.Show
   
    End If
 
End Sub
Private Sub fncMultiply()
  Dim i As Integer
  Randomize
  Me.Caption = "                                      Multiplication  Skill  Test  - Level 3 / " & Z - 6
  Me.Label2.Caption = ""
  Result = 0
  
  
  If Z = 7 Then
    
    Me.txtFirst = Int((9 - 3 + 1) * Rnd + 3)
    Me.txtSecond = Int((8 - 2 + 1) * Rnd + 2)
  
  ElseIf Z = 8 Then
    
    Me.txtFirst = Int((25 - 10 + 1) * Rnd + 10)
    Me.txtSecond = Int((15 - 9 + 1) * Rnd + 9)
  
  ElseIf Z = 9 Then
    
    Me.txtFirst = Int((59 - 21 + 1) * Rnd + 21)
    Me.txtFirst = Int((29 - 15 + 1) * Rnd + 15)
    
  End If
    
    Me.txtResult.Enabled = True
    Me.txtResult.SetFocus
    
    Me.cmdPlay.Enabled = False
    
    Me.Label3.Visible = True
    Me.Label3.Caption = "Time Left : "
    Me.Text4.Visible = True
    Me.Text4 = 10
  
    Me.Timer1.Enabled = True
    Me.Timer1.Interval = 1000
    Me.Timer2.Enabled = False
    varChk = False
    
End Sub
Private Sub cmdQuit_Click()
 Call Form_KeyPress(27)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
 
  If KeyAscii = 27 Then
        Dim Response
        
        Me.Timer1.Enabled = False
        
        Response = MsgBox("Do You Wish To Quit?", vbExclamation + vbYesNo, "Quit !!")
         
        If Response = vbYes Then
           
            sndPlaySound vbNullString, sndAsync
            ScoreBoard
            Unload Me
            
        Else
            Me.Timer1.Enabled = True
        End If
    
    End If
    
End Sub
Private Sub Form_Load()
 
 
 sndPlaySound vbNullString, sndAsync
 
 Me.Timer2.Enabled = False
 Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
 Me.cmdPlay.Picture = LoadPicture(App.Path & "\" & "Child2.jpg")
 Me.cmdPlay.Width = 1575
 
    If Z = 7 Then
        
        Me.Tag = "Multiplication"
        Me.Caption = "                                      Multiplication  Skill  Test  - Level 3 / " & Z - 6
        Me.Picture1.Picture = LoadPicture(App.Path & "\Multiply" & ".bmp")

    ElseIf Z = 10 Then
        
        Me.Tag = "Division"
        Me.Caption = "                                      Division  Skill  Test  - Level 4 / " & Z - 9
        Me.Picture1.Picture = LoadPicture(App.Path & "\Div" & ".bmp")
        
    End If
        
End Sub
Private Sub Form_Unload(Cancel As Integer)
 
 If varPoints > 0 Then
    ScoreBoard
 End If
 
 Unload Me
 
End Sub
Private Sub Timer1_Timer()
 
 Me.Text4 = Text4 - 1
 
 If Me.Text4 = 0 Then
    
    Me.Timer1.Enabled = False
    
    If Val(Me.txtResult) <> Me.txtFirst * Me.txtSecond Then
              
            Me.cmdPlay.Enabled = True
            Me.cmdPlay.SetFocus
            
            Me.txtResult.Enabled = False
            
            Me.Timer1.Enabled = False
            
            Me.Timer2.Enabled = True
            Me.Timer2.Interval = 300
            
            NoisePlay WavMov, SND_ASYNC Or SND_LOOP
            DoEvents
            
            varChk = False
                
            varPoints = varPoints - 500
            Me.txtCurr = -500
            Me.txtTot = Val(Me.txtBal) + Val(Me.txtCurr)
            varPoints = Me.txtTot
        
    End If
        
 End If
 
End Sub
Private Sub Timer2_Timer()
  
  If varChk = True Then
    
    If Z = 10 Or Z = 13 Then
        
        Image2.Picture = LoadPicture(App.Path & "\Level.bmp")
        Image2.Visible = IIf(Image2.Visible = True, False, True)
        
        Image1.Picture = LoadPicture(App.Path & "\animation\A0" & Z1 & ".bmp")
        
        Z1 = Z1 + 1


        If Z1 = 2 Then
            Z1 = 0
        End If
   
   Else
        Image1.Picture = LoadPicture(App.Path & "\animation\B0" & X1 & ".bmp")
        
        X1 = X1 + 1

        If X1 = 4 Then
            X1 = 0
        End If
        
   End If
  
  Else
   
    Image1.Picture = LoadPicture(App.Path & "\animation\L0" & Y1 & ".bmp")
    Y1 = Y1 + 1
    
    If Y1 = 3 Then
        Y1 = 0
    End If
    
  End If
  
End Sub
Private Sub txtResult_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 And Me.txtResult <> "" Then
    Me.cmdPlay.Enabled = True
    Me.cmdPlay.SetFocus
    Me.Timer1.Enabled = False
 End If
 
End Sub
Private Sub txtResult_LostFocus()
  
  If Z <= 9 Then
    
    If Val(Me.txtResult) = Me.txtFirst * Me.txtSecond Then
        
        Me.cmdPlay.Enabled = True
        Me.txtResult.Enabled = False
        varChk = True
        
        X1 = 1
        Z1 = 1
        
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 300
       
        Me.Label3.Visible = True
        Me.Label3.Caption = "Time Left : "
        Me.Text4.Visible = True
    
        Me.txtBal = Val(varPoints)
        varPoints = (Z * 1000) + (Me.Text4 * 125)
        Me.Label6.Caption = "Current Points " & Z * 1000 & " + " & Me.Text4 & " (i.e. Time Left) X 125"
             
        Me.txtCurr = varPoints
        Me.txtTot = Val(Me.txtBal) + Val(Me.txtCurr)
        varPoints = Me.txtTot
       
       If Z = 9 Then
        
        If varPoints > 0 Then Call ScoreBoard
        NoisePlay WavClaps1, SND_ASYNC Or SND_LOOP
        
       Else
        
        NoisePlay WavClaps, SND_ASYNC Or SND_LOOP
        
       End If
        
        Z = Z + 1
    
    Else
    
        Me.cmdPlay.Enabled = True
        Me.txtResult.Enabled = False
        
        Me.Label3.Visible = True
        Me.Label3.Caption = "Time Left : "
        Me.Text4.Visible = True
    
        varChk = False
        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = True
        
        NoisePlay WavMov, SND_ASYNC Or SND_LOOP
    
    End If
 
 ElseIf Z > 9 Then
   Call DivResult
 End If
 
End Sub
Private Sub ScoreBoard()
 On Error Resume Next
  
  sndPlaySound vbNullString, sndAsync
  
  'Exit Sub
  
  ConnectAccessDb
  Set rs = dbs.OpenRecordset("SELECT * FROM tblScore", dbOpenDynaset)
  
  With rs
   
    If rs.BOF Then
     usrId = 1
    
    ElseIf usrId = 0 Then
     
     rs.MoveLast
     usrId = rs!kidid + 1
     
    End If
    
    rs.FindFirst "KidId=" & usrId
     
    If rs.NoMatch = True Then
        
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
   
End Sub
Private Sub fncDivision()
  Dim i As Integer, varRes As Integer
    
  Randomize
  Me.Caption = "                                      Division  Skill  Test  - Level 4 / " & Z - 9
  Me.Label2.Caption = ""
  Me.cmdPlay.Enabled = False
  
  If Z = 10 Then
            
    Me.txtFirst = Int((15 - 2 + 1) * Rnd + 2)
    Me.txtSecond = Int((9 - 2 + 1) * Rnd + 2)
    Me.txtFirst = Me.txtFirst * Me.txtSecond
    
  ElseIf Z = 11 Then
    
    Me.txtFirst = Int((25 - 9 + 1) * Rnd + 9)
    Me.txtSecond = Int((12 - 5 + 1) * Rnd + 5)
    Me.txtFirst = Me.txtFirst * Me.txtSecond
  
  ElseIf Z = 12 Then
    
    Me.txtFirst = Int((30 - 12 + 1) * Rnd + 12)
    Me.txtSecond = Int((18 - 12 + 1) * Rnd + 12)
    Me.txtFirst = Me.txtFirst * Me.txtSecond
    
  End If
    
    Me.txtResult.Enabled = True
    Me.txtResult.SetFocus
    
    Me.Label3.Visible = True
    Me.Label3.Caption = "Time Left : "
    Me.Text4.Visible = True
    Me.Text4 = 10
  
    Me.Timer1.Enabled = True
    Me.Timer1.Interval = 1000
    Me.Timer2.Enabled = False
    varChk = False
    
End Sub
Private Sub DivResult()
 
 If Val(Me.txtResult) = Me.txtFirst / Me.txtSecond Then
        
        Me.cmdPlay.Enabled = True
        Me.txtResult.Enabled = False
        varChk = True
        
        X1 = 1
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 300
        
       
       
        Me.Label3.Visible = True
        Me.Label3.Caption = "Time Left : "
        Me.Text4.Visible = True
    
        Me.txtBal = Val(varPoints)
        varPoints = (Z * 1000) + (Me.Text4 * 125)
        Me.Label6.Caption = "Current Points " & Z * 1000 & " + " & Me.Text4 & " (i.e. Time Left) X 125"
             
        Me.txtCurr = varPoints
        Me.txtTot = Val(Me.txtBal) + Val(Me.txtCurr)
        varPoints = Me.txtTot
        
        If Z = 12 Then
            
            If varPoints > 0 Then Call ScoreBoard
            Me.cmdPlay.Picture = LoadPicture(App.Path & "\" & "level1.bmp")
            NoisePlay WavClaps1, SND_ASYNC Or SND_LOOP

        Else
        
            NoisePlay WavClaps, SND_ASYNC Or SND_LOOP
        
        End If
       
        Z = Z + 1
    
    Else
    
        Me.cmdPlay.Enabled = True
        Me.txtResult.Enabled = False
        
        Me.Label3.Visible = True
        Me.Label3.Caption = "Time Left : "
        Me.Text4.Visible = True
        'Me.Text4 = 10
    
        varChk = False
        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = True
        
        NoisePlay WavMov, SND_ASYNC Or SND_LOOP
    
    End If
        DoEvents
        
End Sub
