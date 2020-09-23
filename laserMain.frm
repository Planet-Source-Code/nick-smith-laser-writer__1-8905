VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Laser Writer"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Origin of Laser"
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   3120
      Width           =   4215
      Begin VB.OptionButton Option2 
         Caption         =   "Bottom Right"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Top Right"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3960
      ScaleHeight     =   1185
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3960
      ScaleHeight     =   1065
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ReDraw"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   2340
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   3
      Top             =   0
      Width           =   3975
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2640
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   2
      Top             =   0
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      MaxLength       =   25
      TabIndex        =   0
      Text            =   "Sample Text"
      Top             =   2325
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Click on the boxes on the right to select laser color and write color.  Click ReDraw to laser write the text."
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Text:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2355
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Counter2 As Long
Private Sub Command1_Click()
Picture1.Cls
Picture2.Cls
Picture1.ForeColor = Picture4.BackColor
TextOut Picture1.hDC, 0, Picture1.Height / 2 - 20, Text1.Text, Len(Text1.Text)
Laser
End Sub

Private Sub Timer1_Timer()
End Sub

Private Sub Laser()
If Option1.Value = True Then
laserPosX& = Picture2.ScaleWidth
laserPosY& = 0
ElseIf Option2.Value = True Then
laserPosX& = Picture2.ScaleWidth
laserPosY& = Picture2.ScaleHeight
End If

For Counter1 = 0 To Picture1.ScaleWidth
    For Counter2 = 0 To Picture1.ScaleHeight
    If GetPixel(Picture1.hDC, Counter1, Counter2) = Picture4.BackColor Then
    Picture2.Line (laserPosX&, laserPosY)-(Counter1, Counter2), Picture3.BackColor
    SetPixel Picture2.hDC, Counter1, Counter2, Picture4.BackColor
    hold (1E-35)
    Picture2.Line (laserPosX&, laserPosY)-(Counter1, Counter2), Picture2.BackColor
    SetPixel Picture2.hDC, Counter1, Counter2, Picture4.BackColor
    End If
    Next Counter2
Next Counter1
End Sub

Private Sub Form_Load()
Option2.Value = True
End Sub


Private Sub Picture3_Click()
CommonDialog1.ShowColor
Picture3.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture4_Click()
CommonDialog1.ShowColor
If CommonDialog1.Color = Picture1.BackColor Then
MsgBox "You cant write with the same color as the backround!", 48, "Error"
Exit Sub
End If
Picture4.BackColor = CommonDialog1.Color
End Sub
