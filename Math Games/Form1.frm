VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "   "
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Height          =   455
      Left            =   6480
      Picture         =   "Form1.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   4920
      Width           =   870
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
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   5280
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
      TabIndex        =   92
      Top             =   4920
      Width           =   1215
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
      TabIndex        =   91
      Top             =   4560
      Width           =   1215
   End
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
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   7080
      TabIndex        =   86
      Top             =   360
      Visible         =   0   'False
      Width           =   855
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
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   82
      Top             =   3720
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
      Height          =   495
      Left            =   6120
      Picture         =   "Form1.frx":2CE4
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1060
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   680
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   280
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00808080&
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
      BackColor       =   &H00FFFFFF&
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
   Begin VB.Image Image2 
      Height          =   645
      Left            =   6240
      Picture         =   "Form1.frx":359F
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
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
      TabIndex        =   90
      Top             =   5280
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080C0FF&
      X1              =   5160
      X2              =   8160
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080C0FF&
      Height          =   5415
      Left            =   5160
      Top             =   240
      Width           =   3015
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
      TabIndex        =   89
      Top             =   4920
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      Height          =   1170
      Left            =   45
      Top             =   4485
      Width           =   5055
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
      TabIndex        =   88
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   2010
      Left            =   5400
      Picture         =   "Form1.frx":7305
      Top             =   840
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
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   5640
      TabIndex        =   87
      Top             =   360
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
      ForeColor       =   &H0080FFFF&
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
      Left            =   5640
      TabIndex        =   83
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00C000C0&
      DrawMode        =   4  'Mask Not Pen
      Height          =   3765
      Left            =   40
      Top             =   240
      Width           =   5060
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varValue As Integer, Result As Double, varChk As Boolean
Dim X As Integer, Y As Integer, Z1 As Integer, chkval As Double
Private Sub cmd1_Click(Index As Integer)
   
   If Me.Tag = "Addition" Then
      
    If Me.cmd1(Index).Caption <> Empty Then
    
        varValue = Val(Me.cmd1(Index).Caption)
        Me.SetFocus
        
        Me.cmd1(Index).Caption = ""
        
    End If
    
      Call fncAdd
   
   Else
    
    If Me.cmd1(Index).Caption <> Empty Then
    
            varValue = Val(Me.cmd1(Index).Caption)
            Me.SetFocus
            Me.cmd1(Index).Caption = ""
    
    End If
    
      Call fncMinus
        
   End If
    
End Sub
Private Sub cmdPlay_Click()
    
    Me.Timer2.Enabled = False
    Me.Image2.Visible = False
    
    sndPlaySound vbNullString, sndAsync
    DoEvents
    
    Me.Image1.Picture = LoadPicture(App.Path & "\Child" & ".jpg")
    
  If Z <= 3 Then
     
     Me.Tag = "Addition"
     Me.txtBal = varPoints
     Me.txtCurr = 0
     Me.txtTot = varPoints + Val(Me.txtCurr)
     Call RepeatGame
  
  ElseIf Z <= 6 Then
     
     Me.Tag = "Subtraction"
     Me.txtBal = varPoints
     Me.txtCurr = 0
     Me.txtTot = varPoints + Val(Me.txtCurr)
     Call NextLevel
     
  Else
     
     Unload Me
     Load Form2
     Form2.Show
     
  End If

End Sub
Private Sub cmdQuit_Click()
 Call Form_KeyPress(27)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Dim Response, retVal
        
        Me.Timer1.Enabled = False
        
        Response = MsgBox("Do You Wish To Quit Game ?", vbExclamation + vbYesNo, "Quit !!")
         
        If Response = vbYes Then
            
            sndPlaySound vbNullString, sndAsync
            
            If varPoints > 0 Then Call ScoreBoard
            Unload Me
            
        Else
           If Me.Text4 > 0 Then Me.Timer1.Enabled = True
        End If
    
    End If
    
End Sub
Private Sub Form_Load()
 
 Dim retVal
 sndPlaySound vbNullString, sndAsync
 
 Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
 
 Me.Timer2.Enabled = False

    If Z = 1 Then
       
       Me.Tag = "Addition"
       Me.Caption = "                                           Addition  Skill  Test  - Level 1 "
    
    ElseIf Z = 4 Then
        
        Me.Tag = "Substraction"
        Me.Caption = "                                           Substraction  Skill  Test  - Level 1 "
        
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub
Private Sub Timer1_Timer()
 Dim retVal
 
 Me.Text4 = Text4 - 1
 
 If Me.Text4 = 0 Then
    
    Me.Timer1.Enabled = False
    
    If Result <> Me.Text2 Then
              
            Me.cmdPlay.Enabled = True
            Me.Timer1.Enabled = False
        
            For i = 0 To cmd1.Count - 1
      
                cmd1(i).Enabled = False
                
            Next i
        
            
            Me.Timer2.Enabled = True
            Me.Timer2.Interval = IIf(Z = 3, 400, 150)
            
            varChk = False
            varPoints = varPoints - 500
            Me.txtCurr = -500
            Me.txtTot = Val(Me.txtBal) + Val(Me.txtCurr)
            varPoints = Me.txtTot
            
            NoisePlay WavMov, SND_ASYNC Or SND_LOOP
            
    End If
        
 End If
 
End Sub
Private Sub Timer2_Timer()
  
  If varChk = True Then
    
    If Z = 4 Or Z = 7 Then
        
        Image1.Picture = LoadPicture(App.Path & "\animation\A0" & Z1 & ".bmp")
        Z1 = Z1 + 1
        
        Image2.Picture = LoadPicture(App.Path & "\Level.bmp")
        Image2.Visible = IIf(Image2.Visible = True, False, True)
        
        If Z1 = 2 Then
            Z1 = 0
        End If
   
   Else
        
        Image1.Picture = LoadPicture(App.Path & "\animation\P0" & X & ".bmp")
        X = X + 1

        If X = 16 Then
            X = 0
        End If
        
   End If
   
  
  Else
   
    Image1.Picture = LoadPicture(App.Path & "\animation\L0" & Y & ".bmp")
    Y = Y + 1
    
    If Y = 3 Then
        Y = 0
    End If
    
  End If
  
End Sub
Private Sub RepeatGame()
  
  Dim i As Integer
  Randomize
  Me.Caption = "                                           Addition  Skill  Test  - Level 1 / " & Z
  Me.Text1 = ""
  Me.Label2.Caption = ""
  Me.Text3 = ""
  Result = 0
  
  For i = 0 To cmd1.Count - 1
      
     cmd1(i).Enabled = True
     Me.cmd1(i).FontSize = 10
     Me.cmd1(i).FontItalic = False
     Me.cmd1(i).FontStrikethru = False
     cmd1(i).Caption = Int((9 * Rnd) + 1)
    
  Next i
  
  Me.cmd1(0).SetFocus
  Me.cmdPlay.Enabled = False
  
  Randomize
  
  Select Case Z
     
     Case 1
        Me.Text2 = Int((30 - 10 + 1) * Rnd + 10)
     Case 2
        Me.Text2 = Int((50 - 25 + 1) * Rnd + 25)
     Case 3
        Me.Text2 = Int((70 - 40 + 1) * Rnd + 40)
  
  End Select
 
  Me.Label1.Visible = True
  Me.Label3.Visible = True
  Me.Label3.Caption = "Time Left : "
  Me.Text4.Visible = True
  Me.Text4 = 10
  
  Me.Timer1.Enabled = True
  Me.Timer1.Interval = 1000
  Me.Timer2.Enabled = False
  varChk = False
 
End Sub
Private Sub NextLevel()
 
 Dim i As Integer
  Randomize
  Me.Caption = "                                           Subtraction  Skill  Test  - Level 2 / " & Z - 3
  Me.Text1 = ""
  Me.Label2.Caption = ""
  Me.Text3 = ""
  Result = 0
  
  For i = 0 To cmd1.Count - 1
      
     cmd1(i).Enabled = True
     Me.cmd1(i).FontSize = 10
     Me.cmd1(i).FontItalic = False
     Me.cmd1(i).FontStrikethru = False
     cmd1(i).Caption = Int((70 * Rnd) + 1)
     
  Next i
  
  Me.cmd1(0).SetFocus
  Me.cmdPlay.Enabled = False
  
  Randomize
  
  Select Case Z
     
     Case 4
        Me.Text2 = Int((29 - 9 + 1) * Rnd + 19)
     Case 5
        Me.Text2 = Int((45 - 22 + 1) * Rnd + 22)
     Case 6
        Me.Text2 = Int((65 - 40 + 1) * Rnd + 40)
  
  End Select
 
  
  Me.Label1.Visible = True
  Me.Label3.Visible = True
  Me.Label3.Caption = "Time Left : "
  Me.Text4.Visible = True
  Me.Text4 = 15
  
  Me.Timer1.Enabled = True
  Me.Timer1.Interval = 1000
  Me.Timer2.Enabled = False
  varChk = False
 
End Sub
Private Sub fncAdd()
  
  Dim rc As Integer, picFile As String
 
        Result = Result + varValue
        Me.Text1 = IIf(Me.Text1 = Empty, varValue, Me.Text1 & " + " & varValue)
        Me.Label2.Caption = "="
        Me.Text3 = Result
    
    If Result = Me.Text2 Then
        
        Me.cmdPlay.Enabled = True
        Me.Timer1.Enabled = False
       
        For i = 0 To cmd1.Count - 1
      
            cmd1(i).Enabled = False
     
        Next i
            
            X = 1
            Me.Timer2.Enabled = True
            Me.Timer2.Interval = IIf(Z = 3, 400, 150)
            
            varChk = True
            
            Me.txtBal = Val(varPoints)
            varPoints = (Z * 1000) + (Me.Text4 * 125)
            Me.Label6.Caption = "Current Points " & Z * 1000 & " + " & Me.Text4 & " (i.e. Time Left) X 125"
             
            Me.txtCurr = varPoints
            Me.txtTot = Val(Me.txtBal) + Val(Me.txtCurr)
            varPoints = Me.txtTot
            
            If Z = 3 Then
               
               If varPoints > 0 Then Call ScoreBoard
               
               NoisePlay WavClaps1, SND_ASYNC Or SND_LOOP
               
            Else
               
               NoisePlay WavClaps, SND_ASYNC Or SND_LOOP
               
            End If
            
            Z = Z + 1
        
    ElseIf Result > Me.Text2 Then
        
            Me.cmdPlay.Enabled = True
            Me.Timer1.Enabled = False
        
            For i = 0 To cmd1.Count - 1
      
                cmd1(i).Enabled = False
     
            Next i
        
            
            Me.Timer2.Enabled = True
            Me.Timer2.Interval = 150
            DoEvents
            
            NoisePlay WavMov, SND_ASYNC Or SND_LOOP
            varChk = False
            
            Me.txtBal = Val(varPoints)
            varPoints = -500
            Me.Label6.Caption = "Current Points                                                     : "
            Me.txtCurr = varPoints
            Me.txtTot = Val(Me.txtBal) + Val(Me.txtCurr)
            varPoints = Me.txtTot
      
    End If
      
End Sub
Private Sub fncMinus()
    
    Dim rc As Integer, picFile As String
    Result = IIf(Result = 0, varValue, Result - varValue)
            
    Me.Text1 = IIf(Me.Text1 = Empty, varValue, Me.Text1 & " - " & varValue)
    Me.Label2.Caption = "="
    Me.Text3 = Result
    
    If Result = Me.Text2 Then
        
        Me.cmdPlay.Enabled = True
        Me.Timer1.Enabled = False
       
        For i = 0 To cmd1.Count - 1
      
            cmd1(i).Enabled = False
     
        Next i
            
            X = 1
            Me.Timer2.Enabled = True
            Me.Timer2.Interval = IIf(Z = 6, 400, 200)
            
            varChk = True
            
            Me.txtBal = Val(varPoints)
            varPoints = (Z * 1000) + (Me.Text4 * 125)
            Me.Label6.Caption = "Current Points " & Z * 1000 & " + " & Me.Text4 & " (i.e. Time Left) X 125"
             
            Me.txtCurr = varPoints
            Me.txtTot = Val(Me.txtBal) + Val(Me.txtCurr)
            varPoints = Me.txtTot
            
            If Z = 6 Then
            
            If varPoints > 0 Then Call ScoreBoard
            
            Me.cmdPlay.Picture = LoadPicture(App.Path & "\" & "level1.bmp")
            NoisePlay WavClaps1, SND_ASYNC Or SND_LOOP
            
           Else
            NoisePlay WavClaps, SND_ASYNC Or SND_LOOP
           End If
           
            Z = Z + 1
        
    ElseIf Result < Me.Text2 Then
        
            Me.cmdPlay.Enabled = True
            Me.Timer1.Enabled = False
        
            For i = 0 To cmd1.Count - 1
      
                cmd1(i).Enabled = False
     
            Next i
        
            
            Me.Timer2.Enabled = True
            Me.Timer2.Interval = 150
            DoEvents
            
            NoisePlay WavMov, SND_ASYNC Or SND_LOOP
            varChk = False
            
            Me.txtBal = Val(varPoints)
            varPoints = -500
            Me.Label6.Caption = "Current Points                                                     : "
            Me.txtCurr = varPoints
            Me.txtTot = Val(Me.txtBal) + Val(Me.txtCurr)
            varPoints = Me.txtTot
      
    End If
        
End Sub
Private Sub ScoreBoard()
 On Error Resume Next
  
  sndPlaySound vbNullString, sndAsync
  
  ConnectAccessDb
  
  Set rs = dbs.OpenRecordset("SELECT * FROM tblScore", dbOpenDynaset)
  
  With rs
    
    If rs.BOF Then
        
        usrId = 1
    
    ElseIf usrId = 0 Then
        
        rs.MoveLast
        usrId = rs!kidId + 1
        
    End If
    
    rs.FindFirst "KidId=" & usrId
      
    If rs.NoMatch = True Then
        
        .AddNew
        !kidId = usrId
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
