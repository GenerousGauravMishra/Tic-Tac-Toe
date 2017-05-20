VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command10"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command10"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command10"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8280
      TabIndex        =   9
      Top             =   1080
      Width           =   2175
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Option2"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Aparajita"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   15
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      TabIndex        =   12
      Top             =   120
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   6495
      Left            =   120
      Top             =   120
      Width           =   7935
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   8040
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   8055
      Y1              =   2280
      Y2              =   2295
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5400
      X2              =   5400
      Y1              =   120
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2760
      X2              =   2760
      Y1              =   120
      Y2              =   6600
   End
   Begin VB.Menu fl 
      Caption         =   "&File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu abt 
      Caption         =   "&About"
      Begin VB.Menu abta 
         Caption         =   "About Auther"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cur As Integer

Private Sub abta_Click()
Form2.Show
End Sub

'*********************************
'******Thinking+Time=Code*********
'******Project on Tic Tac Toe*****
'******Gaurav Mishra`s Creations***
'*********************************

Private Sub Command1_Click()
Module1.mylogic
If (Option1) Then
Command1.Caption = "O"
Command1.Enabled = False
Call mylogic
Option2.Value = vbChecked
ElseIf (Option2) Then
Command1.Caption = "X"
Command1.Enabled = False
Call mylogic
Option1.Value = vbChecked
End If
End Sub

Private Sub Command10_Click()
Module1.ex

End Sub

Private Sub Command11_Click()
Command1.Caption = "1"
Command1.Enabled = True
Command2.Caption = "2"
Command2.Enabled = True
Command3.Caption = "3"
Command3.Enabled = True
Command4.Caption = "4"
Command4.Enabled = True
Command5.Caption = "5"
Command5.Enabled = True
Command6.Caption = "6"
Command6.Enabled = True
Command7.Caption = "7"
Command7.Enabled = True
Command8.Caption = "8"
Command8.Enabled = True
Command9.Caption = "9"
Command9.Enabled = True
Option1.Value = vbChecked
End Sub

Private Sub Command12_Click()
Unload Me
Form3.Show
End Sub

Private Sub Command2_Click()
Module1.mylogic
If (Option1) Then
Command2.Caption = "O"
Command2.Enabled = False
Call mylogic
Option2.Value = vbChecked
ElseIf (Option2) Then
Command2.Caption = "X"
Command2.Enabled = False
Call mylogic
Option1.Value = vbChecked
End If
End Sub

Private Sub Command3_Click()
Module1.mylogic
If (Option1) Then
Command3.Caption = "O"
Command3.Enabled = False
Call mylogic
Option2.Value = vbChecked
ElseIf (Option2) Then
Command3.Caption = "X"
Command3.Enabled = False
Call mylogic
Option1.Value = vbChecked
End If
End Sub

Private Sub Command4_Click()
Module1.mylogic
If (Option1) Then
Command4.Caption = "O"
Command4.Enabled = False
Call mylogic
Option2.Value = vbChecked
ElseIf (Option2) Then
Command4.Caption = "X"
Command4.Enabled = False
Call mylogic
Option1.Value = vbChecked
End If
End Sub

Private Sub Command5_Click()
Module1.mylogic
If (Option1) Then
Command5.Caption = "O"
Command5.Enabled = False
Call mylogic
Option2.Value = vbChecked
ElseIf (Option2) Then
Command5.Caption = "X"
Command5.Enabled = False
Call mylogic
Option1.Value = vbChecked
End If
End Sub

Private Sub Command6_Click()
Module1.mylogic
If (Option1) Then
Command6.Caption = "O"
Command6.Enabled = False
Call mylogic
Option2.Value = vbChecked
ElseIf (Option2) Then
Command6.Caption = "X"
Command6.Enabled = False
Call mylogic
Option1.Value = vbChecked
End If
End Sub

Private Sub Command7_Click()
Module1.mylogic
If (Option1) Then
Command7.Caption = "O"
Command7.Enabled = False
Call mylogic
Option2.Value = vbChecked
ElseIf (Option2) Then
Command7.Caption = "X"
Command7.Enabled = False
Call mylogic
Option1.Value = vbChecked
End If
End Sub

Private Sub Command8_Click()
Module1.mylogic
If (Option1) Then
Command8.Caption = "O"
Command8.Enabled = False
Call mylogic
Option2.Value = vbChecked
ElseIf (Option2) Then
Command8.Caption = "X"
Command8.Enabled = False
Call mylogic
Option1.Value = vbChecked
End If
End Sub

Private Sub Command9_Click()
Module1.mylogic
If (Option1) Then
Command9.Caption = "O"
Command9.Enabled = False
Call mylogic
Option2.Value = vbChecked
ElseIf (Option2) Then
Command9.Caption = "X"
Command9.Enabled = False
Call mylogic
Option1.Value = vbChecked
End If
End Sub

Private Sub Exit_Click()
Module1.ex
End Sub

Private Sub Form_Load()
Option1.Value = vbChecked
Form1.Caption = "Tic-Tac-Toe                                                             (Gaurav Mishra`s Creations)"
Frame1.Caption = "Player"
Frame1.FontBold = True
Option1.Caption = "Player 1"
Option2.Caption = "Player 2"
Option1.FontBold = True
Option2.FontBold = True
Label1.Caption = "Tic-Tac-Toe"
Label2.Caption = "Gaurav Mishra`s Creations.."
Command1.Caption = "1"
Command2.Caption = "2"
Command3.Caption = "3"
Command4.Caption = "4"
Command5.Caption = "5"
Command6.Caption = "6"
Command7.Caption = "7"
Command8.Caption = "8"
Command9.Caption = "9"
Command10.Caption = "EXIT"
Command11.Caption = "Reset"
Command12.Caption = "Change Players"
Option1.Caption = Form3.Text1.Text
Option2.Caption = Form3.Text2.Text

'Form1.BackColor = vbGreen


End Sub


'A Successfull End Of Project.
'Thank You God for giving me the ability to complete this.















