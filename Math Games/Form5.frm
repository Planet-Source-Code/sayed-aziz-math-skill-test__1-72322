VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6780
   ClientLeft      =   2580
   ClientTop       =   7140
   ClientWidth     =   7140
   Icon            =   "Form5.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Left            =   2040
      Top             =   6360
   End
   Begin VB.Timer Timer2 
      Left            =   1440
      Top             =   6360
   End
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   6360
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   3600
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form5.frx":1CCA
      Height          =   3345
      Left            =   160
      TabIndex        =   0
      Top             =   3240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5900
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   8421504
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   19
      RowDividerStyle =   5
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "KDate"
         Caption         =   "  Date / Time"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy h:mm AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "KidName"
         Caption         =   "    Player's Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "KidScore"
         Caption         =   "       Score"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1844.787
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1920.189
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   2055
      Left            =   5160
      Top             =   840
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   2055
      Left            =   2520
      Top             =   840
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   120
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ranker's Scores"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   6735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   3480
      Left            =   105
      Top             =   3180
      Width           =   6870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Math Skill Score Board"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, Y As Integer, Z As Integer, wavClap As String
Private Sub Form_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 27 Then
    Unload Me
 End If
 
End Sub
Private Sub Form_Load()
On Error Resume Next
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    With Adodc1
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
                            
        .ConnectionString = "DSN=MS Access Database;DBQ=" & App.Path & "\Animation\Score.mdb;PWD=zohaib?308;UID=admin;"
        
        .CommandType = adCmdText
        .RecordSource = "SELECT TOP 10 tblscore.KidName,tblscore.KidScore,tblScore.Kdate From tblScore" _
        & " ORDER BY tblscore.KidScore DESC;"
        .Refresh
        
    End With
    
    Me.Timer1.Enabled = True
    Me.Timer2.Enabled = True
    Me.Timer3.Enabled = True
    
    Me.Timer1.Interval = 150
    Me.Timer2.Interval = 400
    Me.Timer3.Interval = 400
    X = 1
    Y = 1
    Z = 1
    
    NoisePlay WavClaps, SND_ASYNC Or SND_LOOP
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
     
    Me.Timer1.Enabled = False
    Me.Timer2.Enabled = False
    Me.Timer3.Enabled = False
    
    Unload Me
    
    Load Form4
    Form4.Show
    varPoints = Empty
    
End Sub
Private Sub Timer1_Timer()
        
        Image2.Picture = LoadPicture(App.Path & "\animation\P0" & X & ".bmp")
        X = X + 1

        If X = 16 Then
            X = 0
        End If
        
End Sub
Private Sub Timer2_Timer()
        
        Image1.Picture = LoadPicture(App.Path & "\animation\A0" & Y & ".bmp")
        Y = Y + 1

        If Y = 2 Then
            Y = 0
        End If
End Sub
Private Sub Timer3_Timer()
        
        Image3.Picture = LoadPicture(App.Path & "\animation\B0" & Z & ".bmp")
        
        Z = Z + 1

        If Z = 4 Then
            Z = 0
        End If
        
End Sub
