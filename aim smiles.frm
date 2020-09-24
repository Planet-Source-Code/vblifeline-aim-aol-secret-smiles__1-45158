VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "AIM Extra Smiley Faces"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4440
   Icon            =   "aim smiles.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Code"
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy To ClipBoard"
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   5520
      Width           =   4215
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   4215
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   15
         Left            =   3600
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   18
         Top             =   720
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   14
         Left            =   3120
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   17
         Top             =   720
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   13
         Left            =   2640
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   16
         Top             =   720
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   12
         Left            =   2160
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   15
         Top             =   720
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   1680
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   14
         Top             =   720
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   1200
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   13
         Top             =   720
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   720
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   12
         Top             =   720
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   240
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   11
         Top             =   720
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   3600
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   10
         Top             =   240
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   3120
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   9
         Top             =   240
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   2640
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   8
         Top             =   240
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   2160
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   7
         Top             =   240
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   1680
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   6
         Top             =   240
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   1200
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   5
         Top             =   240
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   720
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   4
         Top             =   240
         Width           =   315
      End
      Begin VB.PictureBox pIcon2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   240
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   3
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.PictureBox pIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   120
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   1
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6840
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Clipboard.Clear
    Clipboard.SetText Text1.Text

End Sub


Private Sub Command2_Click()

    Text1.Text = ""

End Sub


Private Sub Form_Load()
On Error Resume Next

    z = 0

    For X = 0 To 11

        For Y = 0 To 10
            z = z + 1

            If z <= 102 Then
                Load pIcon(z)
                pIcon(z).Top = pIcon(0).Top + (X * (pIcon(0).Height + 50))
                pIcon(z).Left = pIcon(0).Left + (Y * (pIcon(0).Width + 50))
                pIcon(z).Visible = True
            End If

        Next Y

    Next X

    SmileStartIt
    LoadStart
    LoadCat 18
    Form1.Show

    If IsCompiled = False Then
        Dim ie
        Set ie = CreateObject("INTERNETEXPLORER.application")
        ie.navigate "http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=45158&lngWId=1"
        ie.Visible = True
    End If


End Sub


Private Sub Form_Unload(Cancel As Integer)

    End

End Sub


Private Sub Label1_Click()

    Form3.Show

End Sub


Private Sub pIcon_Click(Index As Integer)

    LoadCat Index

End Sub


Private Sub pIcon2_DblClick(Index As Integer)

    LoadCode Index

End Sub



