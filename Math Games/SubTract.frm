VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   "
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7410
   FillColor       =   &H00E0E0E0&
   Icon            =   "SubTract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   120
      Top             =   4320
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
      Left            =   6480
      TabIndex        =   86
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   4320
   End
   Begin VB.TextBox Text3 
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
      ForeColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   4320
      TabIndex        =   85
      Top             =   4020
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   82
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   130
      Locked          =   -1  'True
      TabIndex        =   81
      Top             =   4020
      Width           =   3840
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Ready"
      Height          =   375
      Left            =   5760
      TabIndex        =   80
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   79
      Left            =   4420
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   78
      Left            =   3800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   77
      Left            =   3180
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   76
      Left            =   2560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   75
      Left            =   1940
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   74
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   73
      Left            =   700
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   72
      Left            =   100
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   71
      Left            =   4420
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   70
      Left            =   3800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   69
      Left            =   3180
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   68
      Left            =   2560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   67
      Left            =   1940
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   66
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   65
      Left            =   700
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   64
      Left            =   100
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   63
      Left            =   4420
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   62
      Left            =   3800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   61
      Left            =   3180
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   60
      Left            =   2560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   59
      Left            =   1940
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   58
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   57
      Left            =   700
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   56
      Left            =   100
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   55
      Left            =   4420
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   54
      Left            =   3800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   53
      Left            =   3180
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   52
      Left            =   2560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   51
      Left            =   1940
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   50
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   49
      Left            =   700
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   48
      Left            =   100
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   47
      Left            =   4420
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   46
      Left            =   3800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   45
      Left            =   3180
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   44
      Left            =   2560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   43
      Left            =   1940
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   42
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   41
      Left            =   700
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   40
      Left            =   100
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   39
      Left            =   4420
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   38
      Left            =   3800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   37
      Left            =   3180
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   36
      Left            =   2560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   35
      Left            =   1940
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   34
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   33
      Left            =   700
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   32
      Left            =   100
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   31
      Left            =   4420
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   30
      Left            =   3800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   29
      Left            =   3180
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   28
      Left            =   2560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   27
      Left            =   1940
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   26
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   25
      Left            =   700
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   24
      Left            =   100
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   23
      Left            =   4420
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1060
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   22
      Left            =   3800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1060
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   21
      Left            =   3180
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1060
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   20
      Left            =   2560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1060
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   19
      Left            =   1940
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1060
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   18
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1060
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   17
      Left            =   700
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1060
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   16
      Left            =   100
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1060
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   15
      Left            =   4420
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   680
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   14
      Left            =   3800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   680
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   13
      Left            =   3180
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   680
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   12
      Left            =   2560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   680
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   11
      Left            =   1940
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   680
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   10
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   680
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   9
      Left            =   700
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   680
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   8
      Left            =   100
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   680
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   7
      Left            =   4420
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   280
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   6
      Left            =   3800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   280
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   5
      Left            =   3180
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   280
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   4
      Left            =   2560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   280
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   3
      Left            =   1940
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   280
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   2
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   280
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   1
      Left            =   700
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   280
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Index           =   0
      Left            =   100
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   280
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   5760
      Top             =   480
      Width           =   1335
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
      Left            =   5160
      TabIndex        =   87
      Top             =   120
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
      TabIndex        =   84
      Top             =   4020
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click Numbers From Left"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   83
      Top             =   3480
      Visible         =   0   'False
      Width           =   2295
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
Dim varValue As Integer, Result As Double, varChk As Boolean, varTot As Integer
Dim wavSetup As String, X As Integer, Y As Integer, Z As Integer
Private Sub cmd1_Click(Index As Integer)
   Dim rc As Integer
   
   If Me.Tag = "Multiplication" Then
        
        If Me.cmd1(Index).Caption <> Empty Then
    
            varValue = Val(Me.cmd1(Index).Caption)
            Me.SetFocus
        
            Me.cmd1(Index).FontStrikethru = True
            Me.cmd1(Index).FontSize = 12
            Me.cmd1(Index).FontName = "Arial"
            Me.cmd1(Index).FontItalic = True
            Me.cmd1(Index).Enabled = False
    
        End If
 
            Result = IIf(Result = 0, varValue, Result * varValue)
            Me.Text1 = IIf(Me.Text1 = Empty, varValue, Me.Text1 & " X " & varValue)
            Me.Label2.Caption = "="
            Me.Text3 = Result
    
        If Result = Me.Text2 Then
        
            Me.cmdPlay.Enabled = True
            Me.Timer1.Enabled = False
       
            For i = 0 To cmd1.Count - 1
      
                cmd1(i).Enabled = False
     
            Next i
        
            'NoisePlay wavSetup, SND_SYNC
        
            Me.Label3 = Empty
            Me.Text4 = Empty
            
            Me.Timer2.Enabled = True
            Me.Timer2.Interval = 20
            DoEvents
            
            wavSetup = NoiseGet(App.Path & "\" & "claps.wav")
            NoisePlay wavSetup, SND_SYNC
            
            varChk = True
            
            Z = Z + 1
        
        ElseIf Result > Me.Text2 Then
        
            Me.cmdPlay.Enabled = True
            Me.Timer1.Enabled = False
        
            For i = 0 To cmd1.Count - 1
      
                cmd1(i).Enabled = False
     
            Next i
        
            'NoisePlay wavSetup, SND_SYNC
            Me.Label3 = Empty
            Me.Text4 = Empty
            Me.Text4 = Empty
            
            Me.Timer2.Enabled = True
            Me.Timer2.Interval = 150
            DoEvents
            
            wavSetup = NoiseGet(App.Path & "\" & "move.wav")
            NoisePlay wavSetup, SND_SYNC
            
            varChk = False
        
        End If
    
    Else
     
      If Me.cmd1(Index).Caption <> Empty Then
    
            varValue = Val(Me.cmd1(Index).Caption)
            Me.SetFocus
        
            Me.cmd1(Index).FontStrikethru = True
            Me.cmd1(Index).FontSize = 12
            Me.cmd1(Index).FontName = "Arial"
            Me.cmd1(Index).FontItalic = True
            Me.cmd1(Index).Enabled = False
    
        End If
 
            Result = IIf(Result = 0, varValue, Result / varValue)
            
            Me.Text1 = IIf(Me.Text1 = Empty, varValue, Me.Text1 & " / " & varValue)
            Me.Label2.Caption = "="
            Me.Text3 = Result
    
        If Result = Me.Text2 Then
        
            Me.cmdPlay.Enabled = True
            Me.Timer1.Enabled = False
       
            For i = 0 To cmd1.Count - 1
      
                cmd1(i).Enabled = False
     
            Next i
        
            'NoisePlay wavSetup, SND_SYNC
        
            Me.Label3 = Empty
            Me.Text4 = Empty
            
            Me.Timer2.Enabled = True
            Me.Timer2.Interval = 20
            DoEvents
            
            wavSetup = NoiseGet(App.Path & "\" & "claps.wav")
            NoisePlay wavSetup, SND_SYNC
            
            varChk = True
            
           Z = Z + 1
        
        ElseIf Result < Me.Text2 Then
        
            Me.cmdPlay.Enabled = True
            Me.Timer1.Enabled = False
        
            For i = 0 To cmd1.Count - 1
      
                cmd1(i).Enabled = False
     
            Next i
        
            'NoisePlay wavSetup, SND_SYNC
            Me.Label3 = Empty
            Me.Text4 = Empty
            Me.Text4 = Empty
            
            Me.Timer2.Enabled = True
            Me.Timer2.Interval = 150
            DoEvents
            
            wavSetup = NoiseGet(App.Path & "\" & "Move.wav")
            NoisePlay wavSetup, SND_SYNC
            
            varChk = False
        
        End If
        
    End If
    
End Sub
Private Sub cmdPlay_Click()
  
  If Z <= 3 Then
     
     Me.Tag = "Multiplication"
     Call RepeatGame
  
  Else
     
     'Z = 4
     Me.Tag = "Division"
     Call NextLevel
     
  End If

End Sub
Private Sub Form_Load()
 
 Me.Caption = "                                           Multiplication  Skill  Test  - Level 1 "
 Me.Timer2.Enabled = False
 Me.Image1.Visible = False
 Me.Tag = "Multiplication"
 Z = 1
 
End Sub
Private Sub Timer1_Timer()
 
 Me.Text4 = Text4 - 1
 
 If Me.Text4 = 0 Then
    
    Me.Timer1.Enabled = False
    
    If Result = Me.Text2 Then
       
       varChk = True
       Timer2.Enabled = True
       Me.Timer2.Interval = 20
       'varTot = varTot + 1
       Z = Z + 1
       
    ElseIf Result <> Me.Text2 Then
              
       Me.Text4 = Empty
       
       For i = 0 To cmd1.Count - 1
      
            cmd1(i).Enabled = False
     
       Next i
       
    End If
        
        varChk = False
        Timer2.Enabled = True
        Me.Timer2.Interval = 150
        Me.cmdPlay.Enabled = True
        
 End If
 
End Sub
Private Sub Timer2_Timer()
  
  Image1.Visible = True
  
  If varChk = True Then
    
    Image1.Picture = LoadPicture(App.Path & "\animation\P0" & X & ".bmp")
    X = X + 1
    'NoisePlay wavSetup, ASND_SYNC
    If X = 16 Then
        X = 0
    End If
  
  Else
   
    Image1.Picture = LoadPicture(App.Path & "\animation\L0" & Y & ".bmp")
    Y = Y + 1
     'NoisePlay wavSetup, ASND_SYNC

    If Y = 2 Then
        Y = 0
    End If
    
  End If
  
End Sub
Private Sub RepeatGame()
  
  Dim i As Integer
  Randomize
  Me.Caption = "                                           Multiplication  Skill  Test  - Level 1 / " & Z
  Me.Text1 = ""
  Me.Label2.Caption = ""
  Me.Text3 = ""
  Result = 0
  
  'wavSetup = NoiseGet(App.Path & "\" & "claps.wav")
  
  For i = 0 To cmd1.Count - 1
      
     cmd1(i).Enabled = True
     Me.cmd1(i).FontSize = 10
     Me.cmd1(i).FontItalic = False
     Me.cmd1(i).FontStrikethru = False
     cmd1(i).Caption = Int(12 * Rnd + 1)
     
  Next i
  
  Me.cmd1(0).SetFocus
  Me.cmdPlay.Enabled = False
  
  Randomize
  
  Select Case Z
     
     Case 1
        
        Randomize
        Me.Text2 = Int((6 - 2 + 1) * Rnd + 2) * Int((6 - 2 + 1) * Rnd + 2)
      
     Case 2
        
        Randomize
        Me.Text2 = Int((8 - 3 + 1) * Rnd + 3) * Int((8 - 3 + 1) * Rnd + 3)
        
     Case 3
        
        Randomize
        Me.Text2 = Int((12 - 4 + 1) * Rnd + 4) * Int((12 - 4 + 1) * Rnd + 4)
        
  End Select
 
  Me.Label1.Visible = True
  Me.Label3.Visible = True
  Me.Label3.Caption = "Time Left : "
  Me.Text4.Visible = True
  Me.Text4 = 10
  
  Me.Timer1.Enabled = True
  Me.Timer1.Interval = 1000
  Me.Image1.Visible = False
  Me.Timer2.Enabled = False
  varChk = False
 
End Sub
Private Sub NextLevel()
 
 Dim i As Integer, TotChk As Integer, TotChk1 As Integer
  Randomize
  Me.Caption = "                                           Division  Skill  Test  - Level 2 / " & Z
  Me.Text1 = ""
  Me.Label2.Caption = ""
  Me.Text3 = ""
  Result = 0
  
  wavSetup = NoiseGet(App.Path & "\" & "claps.wav")
  
  For i = 0 To cmd1.Count - 1
      
     cmd1(i).Enabled = True
     Me.cmd1(i).FontSize = 10
     Me.cmd1(i).FontItalic = False
     Me.cmd1(i).FontStrikethru = False
     cmd1(i).Caption = Int((30 - 5 + 1) * Rnd + 5) * 2
     
  Next i
  
  Me.cmd1(0).SetFocus
  Me.cmdPlay.Enabled = False
   
  Select Case Z
     
     Case 4
        
        Randomize
        TotChk = Int((6 - 2 + 1) * Rnd + 2)
        
        i = 0
        
        For i = 0 To cmd1.Count - 1
                     
            
                
                
        Next i
        
        Me.Text2 = Int((5 - 2 + 1) * Rnd + 2) * 2
      
     Case 5
        
        Randomize
        Me.Text2 = Int((5 - 2 + 1) * Rnd + 2) * 4
        
     Case 6
        
        Randomize
        Me.Text2 = Int((60 - 2 + 1) * Rnd + 2) / Int((60 - 2 + 1) * Rnd + 2)
        Stop
  
  End Select
 
  
  Me.Label1.Visible = True
  Me.Label3.Visible = True
  Me.Label3.Caption = "Time Left : "
  Me.Text4.Visible = True
  Me.Text4 = 10
  
  Me.Timer1.Enabled = True
  Me.Timer1.Interval = 1000
  Me.Image1.Visible = False
  Me.Timer2.Enabled = False
  varChk = False
 
End Sub


