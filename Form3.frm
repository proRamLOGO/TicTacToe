VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12465
   BeginProperty Font 
      Name            =   "Myriad Pro"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   7785
   ScaleWidth      =   12465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command19 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   27
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Play Again"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command11 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   17
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   16
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      MaxLength       =   9
      TabIndex        =   13
      Text            =   "Player 1"
      Top             =   4320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Caption         =   "Timers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   480
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   615
      Begin VB.Timer Timer16 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   120
         Top             =   5640
      End
      Begin VB.Timer Timer15 
         Enabled         =   0   'False
         Interval        =   9000
         Left            =   120
         Top             =   5280
      End
      Begin VB.Timer Timer14 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   120
         Top             =   4920
      End
      Begin VB.Timer Timer13 
         Interval        =   200
         Left            =   120
         Top             =   4560
      End
      Begin VB.Timer Timer12 
         Interval        =   2000
         Left            =   120
         Top             =   4200
      End
      Begin VB.Timer Timer11 
         Interval        =   2000
         Left            =   120
         Top             =   3840
      End
      Begin VB.Timer Timer10 
         Enabled         =   0   'False
         Interval        =   20000
         Left            =   120
         Top             =   3480
      End
      Begin VB.Timer Timer9 
         Enabled         =   0   'False
         Interval        =   50000
         Left            =   120
         Top             =   3120
      End
      Begin VB.Timer Timer8 
         Enabled         =   0   'False
         Interval        =   8000
         Left            =   120
         Top             =   2760
      End
      Begin VB.Timer Timer7 
         Enabled         =   0   'False
         Interval        =   8000
         Left            =   120
         Top             =   2400
      End
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   120
         Top             =   2040
      End
      Begin VB.Timer Timer5 
         Interval        =   1000
         Left            =   120
         Top             =   1680
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   120
         Top             =   1320
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   7000
         Left            =   120
         Top             =   960
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   7000
         Left            =   120
         Top             =   600
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   7000
         Left            =   120
         Top             =   240
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   15
      Top             =   5880
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      MaxLength       =   9
      TabIndex        =   14
      Text            =   "Player 2"
      Top             =   5280
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   3960
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   75.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   75.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   75.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   75.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1440
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   75.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   75.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   75.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2880
         MaskColor       =   &H0000C000&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   75.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   75.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2880
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3000
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   25
      Top             =   6600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Know More"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   20
      Top             =   7080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   19
      Top             =   7080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Let's Play"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   18
      Top             =   6720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image22 
      Height          =   2730
      Left            =   -240
      Picture         =   "Form3.frx":940DC
      Top             =   5880
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image Image10 
      Height          =   270
      Left            =   5040
      Picture         =   "Form3.frx":97414
      Top             =   5760
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image Image9 
      Height          =   1110
      Left            =   3960
      Picture         =   "Form3.frx":976F7
      Top             =   4080
      Visible         =   0   'False
      Width           =   4305
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Has Won"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   5340
      TabIndex        =   26
      Top             =   5505
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      DataMember      =   "&H00E0E0E0&"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1455
      Left            =   5970
      TabIndex        =   24
      Top             =   3960
      Width           =   435
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ! GO !"
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   26.25
         Charset         =   0
         Weight          =   850
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   675
      Left            =   8685
      TabIndex        =   22
      Top             =   7920
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ! GO !"
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   26.25
         Charset         =   0
         Weight          =   850
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   675
      Left            =   1965
      TabIndex        =   21
      Top             =   7920
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Image Image11 
      Height          =   855
      Left            =   4800
      Picture         =   "Form3.frx":992FC
      Top             =   2640
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image Image8 
      Height          =   1020
      Left            =   3240
      Picture         =   "Form3.frx":9A7E3
      Top             =   5280
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image5 
      Height          =   1020
      Left            =   8160
      Picture         =   "Form3.frx":9B36F
      Top             =   5280
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image4 
      Height          =   1095
      Left            =   2400
      Picture         =   "Form3.frx":9BEFB
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   7620
   End
   Begin VB.Image Image13 
      Height          =   975
      Left            =   2280
      Picture         =   "Form3.frx":9F0A2
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.Image Image21 
      Height          =   1575
      Left            =   8880
      Picture         =   "Form3.frx":A1F23
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image19 
      Height          =   1635
      Left            =   8880
      Picture         =   "Form3.frx":A2585
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image18 
      Height          =   1575
      Left            =   2160
      Picture         =   "Form3.frx":A3581
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image20 
      Height          =   1575
      Left            =   2160
      Picture         =   "Form3.frx":A3BE3
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image17 
      Height          =   2730
      Left            =   9840
      Picture         =   "Form3.frx":A4BDF
      Top             =   5880
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   27.75
         Charset         =   0
         Weight          =   850
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   720
      Left            =   6120
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Image16 
      Height          =   6840
      Left            =   9000
      Picture         =   "Form3.frx":A7F16
      Top             =   3240
      Visible         =   0   'False
      Width           =   3330
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   27.75
         Charset         =   0
         Weight          =   850
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   720
      Left            =   6120
      TabIndex        =   10
      Top             =   7920
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Image12 
      BorderStyle     =   1  'Fixed Single
      Height          =   945
      Left            =   0
      Picture         =   "Form3.frx":A95D6
      Stretch         =   -1  'True
      Top             =   7800
      Visible         =   0   'False
      Width           =   12570
   End
   Begin VB.Image Image7 
      Height          =   1050
      Left            =   3240
      Picture         =   "Form3.frx":AF6E3
      Top             =   4320
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image15 
      Height          =   6840
      Left            =   2280
      Picture         =   "Form3.frx":B020A
      Top             =   3360
      Visible         =   0   'False
      Width           =   3330
   End
   Begin VB.Image Image6 
      Height          =   1050
      Left            =   8160
      Picture         =   "Form3.frx":B18CA
      Top             =   4320
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image14 
      Height          =   1710
      Left            =   3600
      Picture         =   "Form3.frx":B23F1
      Top             =   3720
      Visible         =   0   'False
      Width           =   5100
   End
   Begin VB.Image Image23 
      Height          =   2835
      Left            =   2520
      Picture         =   "Form3.frx":B7C3A
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   7365
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   3720
      Picture         =   "Form3.frx":C9A91
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.Image Image3 
      Height          =   2895
      Left            =   2880
      Picture         =   "Form3.frx":D5426
      Top             =   3960
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.Image Image2 
      Height          =   3075
      Left            =   -600
      Picture         =   "Form3.frx":DA4A8
      Top             =   -840
      Width           =   11085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.Enabled = True Then
Form1.MousePointer = 0
Else
Form1.MousePointer = 12
End If
End Sub

Private Sub Command10_Click()
Dim testmsg As Integer
testmsg = MsgBox("Do you want to start the game ?", vbInformation + vbOKCancel, "Start")
If testmsg = 1 Then
Text1.Visible = True
Text2.Visible = True
Image4.Visible = True
Command10.Visible = False
Command11.Visible = False
Command12.Visible = False
Command13.Visible = True
Command19.Visible = True
Image5.Visible = True
Image6.Visible = True
Image7.Visible = True
Image8.Visible = True
Form1.Caption = "Please Enter The Names Of Players"
End If
End Sub

Private Sub Command11_Click()
Command15.Visible = True
Form1.Caption = "About"
Image11.Visible = True
Command14.Visible = True
Image3.Visible = True
Command10.Visible = False
Command11.Visible = False
Command12.Visible = False
End Sub

Private Sub Command12_Click()
Dim testmsg As Integer
testmsg = MsgBox("Do you want to Exit ?", vbCritical + vbOKCancel, "Exit")
If testmsg = 1 Then
Unload Me
Else
Command10.Visible = True
Command11.Visible = True
Command12.Visible = True
End If
End Sub

Private Sub Command13_Click()

If Text1.Text = "Player 1" And Text2.Text = "Player 2" Then
MsgBox "Please enter the Names of Players !", vbCritical, "Names"
Else
Timer14.Enabled = True
Timer9.Enabled = True
Timer8.Enabled = True
Timer7.Enabled = True
Timer5.Enabled = True
Label3.Visible = True
Label2.Visible = True
Image12.Visible = True
Form1.Caption = "Let's Play"
Form1.Visible = True
Timer1.Enabled = True
Timer4.Enabled = True
Timer15.Enabled = True
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
Image4.Visible = False
Image8.Visible = False
Command13.Visible = False
Command19.Visible = False
Text1.Visible = False
Text2.Visible = False
Label10.Caption = Text1.Text
Label11.Caption = Text2.Text
Image9.Visible = True
Image10.Visible = True
End If

End Sub


Private Sub Command1_Click()

If Label11.Visible = True Then
Command1.Caption = "X"
Else
Command1.Caption = "O"
End If

Timer3.Enabled = True

Frame1.Enabled = False
Command1.BackColor = &HC0E0FF
Command1.Enabled = False

End Sub


Private Sub Command14_Click()
Command15.Visible = False
Form1.Caption = "Tic Tac Toe"
Image11.Visible = False
Command14.Visible = False
Image3.Visible = False
Command10.Visible = True
Command11.Visible = True
Command12.Visible = True
End Sub



Private Sub Command18_Click()
Unload Me
End Sub

Private Sub Command19_Click()
Dim testmsg As Integer
testmsg = MsgBox("Do you want to go Back ?", vbQuestion + vbOKCancel, "Back")
If testmsg = 1 Then
Form1.Caption = "Tic Tac Toe"
Command10.Visible = True
Command11.Visible = True
Command12.Visible = True
Command13.Visible = False
Command19.Visible = False
Text1.Visible = False
Text2.Visible = False
Image4.Visible = False
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
Image8.Visible = False
Image9.Visible = False
End If
End Sub

Private Sub Command2_Click()

If Label11.Visible = True Then
Command2.Caption = "X"
Else
Command2.Caption = "O"
End If

Timer3.Enabled = True
Command2.BackColor = &HC0E0FF

Frame1.Enabled = False
Command2.Enabled = False


End Sub


Private Sub Command3_Click()

If Label11.Visible = True Then
Command3.Caption = "X"
Else
Command3.Caption = "O"
End If

Timer3.Enabled = True
Command3.BackColor = &HC0E0FF

Frame1.Enabled = False
Command3.Enabled = False


End Sub

Private Sub Command4_Click()

If Label11.Visible = True Then
Command4.Caption = "X"
Else
Command4.Caption = "O"
End If

Timer3.Enabled = True

Command4.BackColor = &HC0E0FF
Frame1.Enabled = False
Command4.Enabled = False


End Sub

Private Sub Command5_Click()

If Label11.Visible = True Then
Command5.Caption = "X"
Else
Command5.Caption = "O"

End If
Command5.BackColor = &HC0E0FF
Frame1.Enabled = False
Command5.Enabled = False
Timer3.Enabled = True

End Sub

Private Sub Command6_Click()

If Label11.Visible = True Then
Command6.Caption = "X"
Else
Command6.Caption = "O"
End If
Command6.BackColor = &HC0E0FF
Timer3.Enabled = True


Frame1.Enabled = False
Command6.Enabled = False


End Sub

Private Sub Command7_Click()
Command7.BackColor = &HC0E0FF
If Label11.Visible = True Then
Command7.Caption = "X"
Else
Command7.Caption = "O"
End If

Timer3.Enabled = True


Frame1.Enabled = False
Command7.Enabled = False


End Sub

Private Sub Command8_Click()

If Label11.Visible = True Then
Command8.Caption = "X"
Else
Command8.Caption = "O"
End If
Command8.BackColor = &HC0E0FF
Timer3.Enabled = True


Frame1.Enabled = False
Command8.Enabled = False


End Sub

Private Sub Command9_Click()

If Label11.Visible = True Then
Command9.Caption = "X"
Else
Command9.Caption = "O"
End If
Command9.BackColor = &HC0E0FF
Command9.Enabled = False
Timer3.Enabled = True
Frame1.Enabled = False

End Sub





Private Sub Form_Load()
If Label1.Caption = Text1.Text Then
Image5.Visible = True
Image7.Visible = True
Image6.Visible = False
Image8.Visible = False
End If
If Label1.Caption = Text2.Text Then
Image5.Visible = False
Image7.Visible = False
Image6.Visible = True
Image8.Visible = True
End If

End Sub

Private Sub Timer1_Timer()
Label10.Visible = True
Label11.Visible = False
Timer2.Enabled = True
Timer1.Enabled = False
Image19.Visible = True
Image20.Visible = True
Image17.Visible = True
Image22.Visible = True
Timer12.Enabled = False
Image18.Visible = False
Image21.Visible = False
Timer11.Enabled = True

End Sub

Private Sub Timer10_Timer()
If Frame1.Visible = True And Frame1.Enabled = False Then
Image14.Visible = True
Frame1.Visible = False
Image15.Visible = True
Image16.Visible = True
Timer10.Enabled = False
Image1.Visible = False
Image17.Visible = False
Image22.Visible = False
Form1.Height = 8250
Timer6.Enabled = True
Command16.Visible = True
Image19.Visible = False
Image20.Visible = False
Image18.Visible = False
Image21.Visible = False
Command18.Visible = True
End If

If Frame1.Visible = True And Frame1.Enabled = True Then
Image23.Visible = True
Frame1.Visible = False
Image15.Visible = False
Image16.Visible = False
Timer10.Enabled = False
Image1.Visible = False
Image17.Visible = False
Image22.Visible = False
Form1.Height = 8250
Timer6.Enabled = True
Command16.Visible = True
Image19.Visible = False
Image20.Visible = False
Image18.Visible = False
Image21.Visible = False
Command18.Visible = True
End If


End Sub

Private Sub Timer11_Timer()
Image17.Visible = False
Image22.Visible = False
End Sub

Private Sub Timer12_Timer()
Image17.Visible = False
Image22.Visible = False
End Sub

Private Sub Timer13_Timer()
If Image14.Visible = True Or Image23.Visible = True Then
Image17.Visible = False
Image18.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
End If

End Sub

Private Sub Timer14_Timer()
Label11.Visible = False
End Sub



Private Sub Timer15_Timer()
Timer14.Enabled = False
Timer15.Enabled = False
End Sub

Private Sub Timer2_Timer()
Label11.Visible = True
Label10.Visible = False
Image19.Visible = False
Image20.Visible = False
Image18.Visible = True
Image21.Visible = True
Timer1.Enabled = True
Image17.Visible = True
Image22.Visible = True
Timer2.Enabled = False
Timer12.Enabled = True
Timer11.Enabled = False

End Sub

Private Sub Timer3_Timer()
Frame1.Enabled = True
End Sub

Private Sub Timer4_Timer()
Image9.Visible = False
Image10.Visible = False

Frame1.Visible = True
Form1.Height = 9150
Image1.Visible = True
Timer4.Enabled = False
End Sub

Private Sub Command16_Click()

Dim testmsg As Integer
testmsg = MsgBox("Do you want to start the game again ?", vbInformation + vbOKCancel, "Start Again")
If testmsg = 1 Then
Image23.Visible = False
Image14.Visible = False
Image16.Visible = False
Text1.Visible = True
Label1.Visible = False
Label10.Visible = False
Command19.Visible = True
Label4.Visible = False
Text2.Visible = True
Text1.Text = "Player 1"
Text2.Text = "Player 2"
Image4.Visible = True
Command16.Visible = False
Command18.Visible = False
Command13.Visible = True
Image5.Visible = True
Image6.Visible = True
Image15.Visible = False
Image7.Visible = True
Image8.Visible = True
Image13.Visible = False
Form1.Caption = "Please Enter The Names Of Players"
Command1.Caption = ""
Command1.Enabled = True
Command1.BackColor = &H80C0FF
Command2.Caption = ""
Command2.Enabled = True
Command2.BackColor = &H80C0FF
Command3.BackColor = &H80C0FF
Command3.Caption = ""
Command3.Enabled = True
Command4.Caption = ""
Command4.Enabled = True
Command4.BackColor = &H80C0FF
Command5.BackColor = &H80C0FF
Command5.Caption = ""
Command5.Enabled = True
Command6.Caption = ""
Command6.Enabled = True
Command6.BackColor = &H80C0FF
Command7.BackColor = &H80C0FF
Command7.Caption = ""
Command7.Enabled = True
Command8.Caption = ""
Command8.Enabled = True
Command8.BackColor = &H80C0FF
Command9.BackColor = &H80C0FF
Command9.Caption = ""
Command9.Enabled = True
Image19.Visible = False
Image20.Visible = False
Image18.Visible = False
Image21.Visible = False

End If


End Sub

Private Sub Timer5_Timer()

If Command1.Caption = "O" And Command2.Caption = "O" And Command3.Caption = "O" Then
Image13.Visible = True
Command16.Visible = True
Command18.Visible = True
Label1.Caption = Text1.Text
Form1.Height = 8250
Label1.Visible = True
Timer6.Enabled = True
Label4.Visible = True
Image1.Visible = False
Frame1.Visible = False
End If
If Command1.Caption = "O" And Command5.Caption = "O" And Command9.Caption = "O" Then
Image13.Visible = True
Command16.Visible = True
Command18.Visible = True
Label1.Caption = Text1.Text
Form1.Height = 8250
Image1.Visible = False
Frame1.Visible = False
Label1.Visible = True
Timer6.Enabled = True
Label4.Visible = True
End If
If Command1.Caption = "O" And Command4.Caption = "O" And Command7.Caption = "O" Then
Image13.Visible = True
Image1.Visible = False
Frame1.Visible = False
Command16.Visible = True
Command18.Visible = True
Label1.Caption = Text1.Text
Form1.Height = 8250
Label1.Visible = True
Timer6.Enabled = True
Label4.Visible = True
End If
If Command7.Caption = "O" And Command8.Caption = "O" And Command9.Caption = "O" Then
Image13.Visible = True
Command16.Visible = True
Command18.Visible = True
Label1.Caption = Text1.Text
Form1.Height = 8250
Label1.Visible = True
Image1.Visible = False
Frame1.Visible = False
Timer6.Enabled = True
Label4.Visible = True
End If
If Command3.Caption = "O" And Command6.Caption = "O" And Command9.Caption = "O" Then
Image13.Visible = True
Label4.Visible = True
Command16.Visible = True
Command18.Visible = True
Image1.Visible = False
Frame1.Visible = False
Label1.Caption = Text1.Text
Form1.Height = 8250
Label1.Visible = True
Timer6.Enabled = True
End If
If Command4.Caption = "O" And Command5.Caption = "O" And Command6.Caption = "O" Then
Image13.Visible = True
Command16.Visible = True
Command18.Visible = True
Image1.Visible = False
Frame1.Visible = False
Label1.Caption = Text1.Text
Form1.Height = 8250
Label4.Visible = True
Label1.Visible = True
Timer6.Enabled = True
End If
If Command7.Caption = "O" And Command5.Caption = "O" And Command3.Caption = "O" Then
Image13.Visible = True
Label4.Visible = True
Command16.Visible = True
Command18.Visible = True
Label1.Caption = Text1.Text
Form1.Height = 8250
Label1.Visible = True
Timer6.Enabled = True
Image1.Visible = False
Frame1.Visible = False
End If
If Command2.Caption = "O" And Command5.Caption = "O" And Command8.Caption = "O" Then
Image13.Visible = True
Command16.Visible = True
Label4.Visible = True
Command18.Visible = True
Label1.Caption = Text1.Text
Form1.Height = 8250
Label1.Visible = True
Image1.Visible = False
Frame1.Visible = False
Timer6.Enabled = True
End If
If Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X" Then
Image13.Visible = True
Command16.Visible = True
Label4.Visible = True
Command18.Visible = True
Label1.Caption = Text2.Text
Form1.Height = 8250
Image1.Visible = False
Frame1.Visible = False
Label1.Visible = True
Timer6.Enabled = True
End If
If Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X" Then
Image13.Visible = True
Image1.Visible = False
Frame1.Visible = False
Command16.Visible = True
Command18.Visible = True
Label1.Caption = Text2.Text
Form1.Height = 8250
Label1.Visible = True
Timer6.Enabled = True
Label4.Visible = True
End If
If Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X" Then
Image13.Visible = True
Label4.Visible = True
Command16.Visible = True
Command18.Visible = True
Label1.Caption = Text2.Text
Form1.Height = 8250
Label4.Visible = True
Label1.Visible = True
Image1.Visible = False
Frame1.Visible = False
Timer6.Enabled = True
End If
If Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X" Then
Image13.Visible = True
Image1.Visible = False
Frame1.Visible = False
Label4.Visible = True
Command16.Visible = True
Command18.Visible = True
Label1.Caption = Text2.Text
Form1.Height = 8250
Label1.Visible = True
Timer6.Enabled = True
End If
If Command3.Caption = "X" And Command6.Caption = "X" And Command9.Caption = "X" Then
Image13.Visible = True
Label4.Visible = True
Command16.Visible = True
Image1.Visible = False
Frame1.Visible = False
Command18.Visible = True
Label1.Caption = Text2.Text
Form1.Height = 8250
Label1.Visible = True
Timer6.Enabled = True
End If
If Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X" Then
Image13.Visible = True
Label4.Visible = True
Command16.Visible = True
Command18.Visible = True
Image1.Visible = False
Frame1.Visible = False
Label1.Caption = Text2.Text
Form1.Height = 8250
Label1.Visible = True
Timer6.Enabled = True
End If
If Command7.Caption = "X" And Command5.Caption = "X" And Command3.Caption = "X" Then
Image13.Visible = True
Label4.Visible = True
Command16.Visible = True
Command18.Visible = True
Label1.Caption = Text2.Text
Form1.Height = 8250
Label1.Visible = True
Image1.Visible = False
Frame1.Visible = False
Timer6.Enabled = True
End If
If Command2.Caption = "X" And Command5.Caption = "X" And Command8.Caption = "X" Then
Image13.Visible = True
Image1.Visible = False
Frame1.Visible = False
Label4.Visible = True
Command16.Visible = True
Command18.Visible = True
Label1.Caption = Text2.Text
Form1.Height = 8250
Label1.Visible = True
Timer6.Enabled = True
End If




End Sub

Private Sub Timer6_Timer()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False

End Sub

Private Sub Winner_Click()

End Sub







Private Sub Timer8_Timer()
If Label10.Visible = True Then
Frame1.Enabled = True
End If
Timer8.Enabled = False
End Sub

Private Sub Timer9_Timer()
Timer10.Enabled = True
Timer9.Enabled = False
End Sub
