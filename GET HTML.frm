VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Richard's HTML stealer"
   ClientHeight    =   3945
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GET CODE"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   8655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Enter URL here"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Menu mnu1 
      Caption         =   "File"
      Begin VB.Menu mnu2 
         Caption         =   "Save As &HTML File"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu3 
         Caption         =   "Save As &Text File"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnu10 
         Caption         =   "Save As &Frm File"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnu4 
         Caption         =   "Help"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnu5 
         Caption         =   "About HTML Stealer"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu6 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnu7 
      Caption         =   "Design"
      Begin VB.Menu mnu8 
         Caption         =   "Webpage"
         Shortcut        =   ^W
         Tag             =   "up"
      End
      Begin VB.Menu mnu9 
         Caption         =   "Executable"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strURL As String
Dim yourtext As String

Private Sub Command1_Click()
strURL = Text1.Text
If Len(Text1.Text + strURL) > 11 Then
Text2.Text = Inet1.OpenURL(strURL)
ElseIf Len(Text1.Text + strURL) < 11 Then
MsgBox "Please enter a valid URL"
End If
End Sub



Private Sub Command2_Click()
CommonDialog1.DefaultExt = "HTM"
CommonDialog1.Filter = "HTML files (*.HTML;*.HTM)|*.HTML;HTM"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, Text2.Text
    Close #1
End If
End Sub



Private Sub Form_Load()
yourtext = Text2.Text
End Sub

Private Sub mnu10_Click()
CommonDialog1.DefaultExt = "Frm"
CommonDialog1.Filter = "Form Files(*.Frm)|*.Frm"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
Open CommonDialog1.FileName For Output As #1
Print #1, Text2.Text
Close #1
End If
End Sub

Private Sub mnu2_Click()
CommonDialog1.DefaultExt = "HTM"
CommonDialog1.Filter = "HTML files (*.HTML;*.HTM)|*.HTML;HTM"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, Text2.Text
    Close #1
    End If
End Sub

Private Sub mnu3_Click()
CommonDialog1.DefaultExt = "TXT"
CommonDialog1.Filter = "Text Files (*.Txt)|*.Txt"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, Text2.Text & yourtext
    Close #1
    End If
End Sub

Private Sub mnu4_Click()
Form2.Show
End Sub

Private Sub mnu5_Click()
Form3.Show
End Sub
 
 Private Sub mnu6_Click()
 Unload Me
 End Sub

Private Sub mnu8_Click()
If mnu8.Tag = "up" Then
Text2.Text = "<HTML>" & "<HEAD>" & "<TITLE>" & "</TITLE>" & "</HEAD>" & "<BODY>" & "</BODY>" & "</HTML>"
End If
End Sub

Private Sub mnu9_Click()
MsgBox "You can type out the code here, but won't be able to run it directly from here. You will have to open it after saving in your VB or C++ etc program", vbOKOnly
End Sub
