VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fizz Tool Cyber Drive Manager"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "REFRESH"
      Height          =   1095
      Left            =   3720
      TabIndex        =   10
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create Virtual Drive"
      Height          =   1695
      Left            =   1440
      TabIndex        =   8
      Top             =   4680
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":0442
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Form1.frx":0447
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove Virtual Drive"
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   3000
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.DriveListBox drvList 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.DirListBox dirList 
      BackColor       =   &H00FFFFFF&
      Height          =   1890
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Drives will stay connected until they are manually removed!"
      Height          =   735
      Left            =   1200
      TabIndex        =   11
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "The Fizz Tool Cyber Drive Manager is a part of Fizz Tool XP SE.  For more information, visit www.fizzworld.org"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      TabIndex        =   9
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   3360
      X2              =   3360
      Y1              =   0
      Y2              =   6840
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   360
      X2              =   600
      Y1              =   4200
      Y2              =   4440
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   840
      X2              =   600
      Y1              =   4200
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   600
      X2              =   600
      Y1              =   3720
      Y2              =   4440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Virtual Drives"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim b As String
Dim c As String
Dim tempdir As String
Dim temp1dir As String

Private Sub Command1_Click()

Shell ("command.com /c subst/d " & Drive1.Drive)

Call RemoveFromRegistry(Drive1.Drive)



Call Update


End Sub

Private Sub Command2_Click()
If Right$(tempdir, 1) = "\" Then temp1dir = Left$(tempdir, Len(tempdir) - 1)
makeme = LCase$(Right$(List1.Text, 2)) & " " & Chr$(34) & temp1dir & Chr$(34)
Shell ("command.com /c subst " & makeme)

Call WriteRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", LCase$(Right$(List1.Text, 2)), "subst " + makeme)


Call Update


End Sub

Private Sub Command3_Click()
Call Update

End Sub

Private Sub Drive1_Change()
Call Update

End Sub

Private Sub Form_Load()
Call Update
List1.AddItem ("E:")
List1.AddItem ("F:")
List1.AddItem ("G:")
List1.AddItem ("H:")
List1.AddItem ("I:")
List1.AddItem ("J:")
List1.AddItem ("K:")
List1.AddItem ("L:")
List1.AddItem ("M:")
List1.AddItem ("N:")
List1.AddItem ("O:")
List1.AddItem ("P:")
List1.AddItem ("Q:")
List1.AddItem ("R:")
List1.AddItem ("S:")
List1.AddItem ("T:")
List1.AddItem ("U:")
List1.AddItem ("V:")
List1.AddItem ("W:")
List1.AddItem ("X:")
List1.AddItem ("Y:")
List1.AddItem ("Z:")


End Sub

Private Sub DirList_Change()
tempdir = dirList.Path
If Right$(tempdir, 1) <> "\" Then tempdir = tempdir & "\"
Text2.Text = tempdir
    
End Sub


Private Sub drvList_Change()
    On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    Exit Sub
DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub

tempdir = dirList.Path
If Right$(tempdir, 1) <> "\" Then tempdir = tempdir & "\"

Text2.Text = tempdir



End Sub

Private Sub Update()




Text1.Text = ""
b = ""
Shell ("command.com /c subst > temp.blh")
Open "temp.blh" For Input As #1
While Not EOF(1)
Input #1, a
b = b & vbCrLf & a
Wend
Close #1
Text1.Text = b
c = Drive1.Drive

tempdir = dirList.Path
If Right$(tempdir, 1) <> "\" Then tempdir = tempdir & "\"
Text2.Text = tempdir


drvList.Refresh
dirList.Refresh
Drive1.Refresh


End Sub


Private Sub List1_Click()
Call Update

End Sub
