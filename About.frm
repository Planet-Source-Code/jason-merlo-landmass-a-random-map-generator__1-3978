VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About LandMass..."
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Flexgrid experiment number two.  Now we get fancy.  :)"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Label Label2 
      Caption         =   "Flexgrid experiment number two.  Now we get fancy.  :)"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Flexgrid experiment number two.  Now we get fancy.  :)"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

About.Hide

End Sub

Private Sub Form_Load()

Label1.Caption = "LandMass was originally written by me, Jason Merlo, as an MSFlexGrid experiment.  As it turns out, Flexgrids are really, really slow and are not built for graphics."
Label2.Caption = "I then dropped the project for a while but picked it back up once I learned how to use the BitBlt 32-bit API function.  Now tile-based game engines are relatively easy to build."
Label3.Caption = "I intend to turn LandMass into a Risk-style wargame, eventually."


End Sub

