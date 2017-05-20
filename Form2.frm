VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8535
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   6480
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   1755
      TabIndex        =   9
      Top             =   3960
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   240
      Picture         =   "Form2.frx":2DE9
      ScaleHeight     =   2115
      ScaleWidth      =   1755
      TabIndex        =   8
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Label Label8 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3240
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   7935
   End
   Begin VB.Label Label5 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   7935
   End
   Begin VB.Label Label4 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   7935
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   7935
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form2.Caption = "About Auther"
Form2.BackColor = vbWhite
Label1.BackColor = vbWhite
Label2.BackColor = vbWhite
Label3.BackColor = vbWhite
Label4.BackColor = vbWhite
Label5.BackColor = vbWhite
Label6.BackColor = vbWhite
Label2.ForeColor = vbBlue
Label4.ForeColor = vbBlue
Label7.BackColor = vbWhite
Label8.BackColor = vbWhite
Label1.Alignment = 2
Label2.Alignment = 2
Label3.Alignment = 2
Label4.Alignment = 2
Label5.Alignment = 2
Label6.Alignment = 2
Label7.Alignment = 2
Label8.Alignment = 2
Label9.Alignment = 2
Label10.Alignment = 2
Label9.BackColor = vbWhite
Label10.BackColor = vbWhite
Label1.Caption = "Under The Supervision Of"
Label2.Caption = "Mr.Pankaj Kotiyal Sir."
Label9.Caption = "Mr.Pankaj Kotiyal Sir ->"
Label10.Caption = "<- Gaurav Mishra"
Label3.Caption = "Developed By"
Label4.Caption = "Gaurav Mishra"
Label5.Caption = "Dev Bhoomi Group Of Institutions Dehradun."
Label6.Caption = "Co-+919456510074"
Label7.Caption = "Facebook.com/Thestatementmissing"
Label8.Caption = "Thinkprogrammer.blogspot.in"

End Sub
