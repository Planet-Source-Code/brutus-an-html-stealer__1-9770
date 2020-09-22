VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   LinkTopic       =   "Form4"
   ScaleHeight     =   3930
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "*Make Window Applications that can be compiled through your Visual Basic compiler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "*Edit and Create Webpages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "*Steal code quickly and easily"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "HTML STEALER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
Form1.Show
Unload Me
End Sub
