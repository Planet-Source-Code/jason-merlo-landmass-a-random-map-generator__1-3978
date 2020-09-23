VERSION 5.00
Begin VB.Form Land 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LandMass"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10455
   FillStyle       =   0  'Solid
   Icon            =   "Land.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Selection 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7440
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton BorderBut 
      Caption         =   "Bordering"
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   2640
      Width           =   855
   End
   Begin VB.PictureBox Ocean 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   7140
      Left            =   9240
      Picture         =   "Land.frx":0442
      ScaleHeight     =   7080
      ScaleWidth      =   9720
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   9780
   End
   Begin VB.PictureBox LandMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3180
      Left            =   0
      Picture         =   "Land.frx":4B344
      ScaleHeight     =   3120
      ScaleWidth      =   10560
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   10620
   End
   Begin VB.CommandButton RedrawBut 
      Caption         =   "Redraw"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton GoBut 
      Caption         =   "Go!"
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin VB.Label SelectionText 
      Caption         =   "Selection:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Commentary 
      Caption         =   "Running Commentary Text."
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   5040
      Width           =   7695
   End
   Begin VB.Menu MenuFile 
      Caption         =   "File"
      Begin VB.Menu MenuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MenuOptions 
      Caption         =   "Terraform"
      Begin VB.Menu MenuCountrySize 
         Caption         =   "Country Size"
         Begin VB.Menu MenuSize 
            Caption         =   "Tiny Independently-Owned Countries"
            Index           =   1
         End
         Begin VB.Menu MenuSize 
            Caption         =   "Small Third-World Countries"
            Index           =   2
         End
         Begin VB.Menu MenuSize 
            Caption         =   "Medium OPEC Countries"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu MenuSize 
            Caption         =   "Large First-World Countries"
            Index           =   4
         End
         Begin VB.Menu MenuSize 
            Caption         =   "Continents"
            Index           =   5
         End
      End
      Begin VB.Menu MenuCountrySizeProp 
         Caption         =   "Country Size Proportions"
         Begin VB.Menu MenuProp 
            Caption         =   "Proportional Countries"
            Index           =   1
         End
         Begin VB.Menu MenuProp 
            Caption         =   "Somewhat Proportional Countries"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MenuProp 
            Caption         =   "Unproportional Countries"
            Index           =   3
         End
      End
      Begin VB.Menu MenuCountryShapes 
         Caption         =   "Country Shapes"
         Begin VB.Menu MenuShape 
            Caption         =   "Normal (no artificial distortion)"
            Index           =   1
         End
         Begin VB.Menu MenuShape 
            Caption         =   "Irregular"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MenuShape 
            Caption         =   "Very Irregular"
            Index           =   3
         End
      End
      Begin VB.Menu MenuMinLakeSize 
         Caption         =   "Minimum Allowed Lake Size"
         Begin VB.Menu MenuLakeSize 
            Caption         =   "No Lake Correction"
            Index           =   1
         End
         Begin VB.Menu MenuLakeSize 
            Caption         =   "Tiny Lakes"
            Index           =   2
         End
         Begin VB.Menu MenuLakeSize 
            Caption         =   "Medium Lakes"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu MenuLakeSize 
            Caption         =   "Large Lakes"
            Index           =   4
         End
      End
      Begin VB.Menu MenuGlobal 
         Caption         =   "Approximate Land:Water ratio"
         Begin VB.Menu MenuPct 
            Caption         =   "1:9"
            Index           =   1
         End
         Begin VB.Menu MenuPct 
            Caption         =   "1:3"
            Index           =   2
         End
         Begin VB.Menu MenuPct 
            Caption         =   "1:1"
            Index           =   3
         End
         Begin VB.Menu MenuPct 
            Caption         =   "3:1"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu MenuPct 
            Caption         =   "9:1"
            Index           =   5
         End
      End
      Begin VB.Menu MenuIsle 
         Caption         =   "Islands"
         Begin VB.Menu MenuIslands 
            Caption         =   "No Islands"
            Index           =   1
         End
         Begin VB.Menu MenuIslands 
            Caption         =   "Some Islands"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MenuIslands 
            Caption         =   "Lots of Islands"
            Index           =   3
         End
      End
   End
   Begin VB.Menu MenuDraw 
      Caption         =   "Draw"
      Begin VB.Menu MenuBorderChoice 
         Caption         =   "Borders"
         Begin VB.Menu MenuBorders 
            Caption         =   "Big Lines"
            Index           =   1
         End
         Begin VB.Menu MenuBorders 
            Caption         =   "Big Dots"
            Index           =   2
         End
         Begin VB.Menu MenuBorders 
            Caption         =   "Hash Marks"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu MenuBorders 
            Caption         =   "Small Dots"
            Index           =   4
         End
         Begin VB.Menu MenuBorders 
            Caption         =   "None"
            Index           =   5
         End
      End
      Begin VB.Menu MenuColors 
         Caption         =   "Colors"
         Begin VB.Menu MenuColor 
            Caption         =   "Earth Tones"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu MenuColor 
            Caption         =   "Black && White"
            Index           =   2
         End
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "Help"
      Begin VB.Menu MenuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu MenuTextFile 
         Caption         =   "View Text File"
      End
   End
End
Attribute VB_Name = "Land"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'LandMass generator by Jason Merlo  10/10/99
'jmerlo@austin.rr.com
'jason.merlo@frco.com
'http://home.austin.rr.com/smozzie

Option Explicit
Option Base 1

Dim i As Long

Dim MyMap As Map

Dim CurrentMouse As Long

Dim RollOver As Boolean

Dim MinLakeSize As Integer
Dim MaxCountrySize As Integer
Dim NumCountries As Integer
Dim CFGCountrySize As Integer
Dim CFGCheckedPct As Integer
Dim CFGIslands As Integer
Dim CFGLakeSize As Integer
Dim CFGCoast As Integer
Dim CFGProportion As Integer
Dim CFGShape As Integer
Dim CFGColor As Integer
Dim CFGBorders As Integer

Dim LandPct As Double
Dim PropPct As Double
Dim ShapePct As Double
Dim CoastPctKeep As Double
Dim IslePctKeep As Double

Private Sub GoBut_Click()

GoBut.Enabled = False
RedrawBut.Enabled = False
SelectionText.Enabled = False
Selection.Text = ""
Selection.Enabled = False
BorderBut.Enabled = False
RollOver = False

Set MyMap = New Map

Commentary.Caption = "Creating map..."
DoEvents

'Collect the menu settings.
Call CollectMenuSettings

'Clear the previous display.
Call DrawBkg

'This is the syntax for building a map.
MyMap.CreateMap NumCountries, MaxCountrySize, MinLakeSize, LandPct, PropPct, _
     ShapePct, CoastPctKeep, IslePctKeep, CFGColor, CFGIslands

'This command draws the map.
MyMap.DisplayMap LandMap.hDC, Land.hDC, CFGBorders

GoBut.Enabled = True
RedrawBut.Enabled = True
SelectionText.Enabled = True
Selection.Enabled = True
RollOver = True

Commentary.Caption = "LandMass is finished. " & MyMap.NumCountries & " countries present on this map."

End Sub

Private Sub MenuAbout_Click()

About.Show vbModal

End Sub

Private Sub MenuExit_Click()

End

End Sub

Private Sub Form_Load()

Dim WinWidth As Long
Dim WinHeight As Long

Randomize

WinWidth = 750 * Screen.TwipsPerPixelX
WinHeight = 550 * Screen.TwipsPerPixelY

'Set up window position and dimensions.
Land.Width = WinWidth
Land.Height = WinHeight
Land.ScaleWidth = WinWidth
Land.ScaleHeight = WinHeight
Land.ScaleMode = 1

DoEvents

'Add other controls.
GoBut.Left = Land.ScaleWidth - (70 * Screen.TwipsPerPixelX)
GoBut.Top = (10 * Screen.TwipsPerPixelY)
RedrawBut.Left = Land.ScaleWidth - (70 * Screen.TwipsPerPixelX)
RedrawBut.Top = Land.ScaleHeight - (100 * Screen.TwipsPerPixelY)

SelectionText.Left = Land.ScaleWidth - (70 * Screen.TwipsPerPixelX)
SelectionText.Top = 110 * Screen.TwipsPerPixelY
Selection.Left = Land.ScaleWidth - (70 * Screen.TwipsPerPixelX)
Selection.Top = 125 * Screen.TwipsPerPixelY
BorderBut.Left = Land.ScaleWidth - (70 * Screen.TwipsPerPixelX)
BorderBut.Top = 150 * Screen.TwipsPerPixelY

SelectionText.Enabled = False
Selection.Enabled = False
BorderBut.Enabled = False
GoBut.Enabled = True
RedrawBut.Enabled = False
RollOver = False

Commentary.Left = Land.Left + (10 * Screen.TwipsPerPixelX)
Commentary.Top = Land.ScaleHeight - (18 * Screen.TwipsPerPixelY)
Commentary.Caption = "LandMass is idle."

End Sub

Private Sub MenuSize_Click(Index As Integer)

For i = 1 To 5
  MenuSize(i).Checked = False
Next i
MenuSize(Index).Checked = True

End Sub

Private Sub MenuLakeSize_Click(Index As Integer)

For i = 1 To 4
  MenuLakeSize(i).Checked = False
Next i
MenuLakeSize(Index).Checked = True

End Sub

Private Sub MenuPct_Click(Index As Integer)

For i = 1 To 5
  MenuPct(i).Checked = False
Next i
MenuPct(Index).Checked = True

End Sub

Private Sub MenuIslands_Click(Index As Integer)

For i = 1 To 3
  MenuIslands(i).Checked = False
Next i
MenuIslands(Index).Checked = True

End Sub

Private Sub MenuShape_Click(Index As Integer)

For i = 1 To 3
  MenuShape(i).Checked = False
Next i
MenuShape(Index).Checked = True

End Sub

Private Sub MenuProp_Click(Index As Integer)

For i = 1 To 3
  MenuProp(i).Checked = False
Next i
MenuProp(Index).Checked = True

End Sub

Private Sub MenuBorders_Click(Index As Integer)

For i = 1 To 5
  MenuBorders(i).Checked = False
Next i
MenuBorders(Index).Checked = True

End Sub

Private Sub MenuColor_Click(Index As Integer)

For i = 1 To 2
  MenuColor(i).Checked = False
Next i
MenuColor(Index).Checked = True

End Sub

Private Sub MenuTextFile_click()

Shell "Notepad.exe LandMass.txt", vbNormalFocus

End Sub

Private Sub RedrawBut_Click()

Dim l As Integer

'Grab the menu settings, as some of these are cosmetic.
Call CollectMenuSettings

'Set up the CountryColor array.
Select Case CFGColor
  Case 1:  For l = 1 To NumCountries
             MyMap.CountryColor(l) = l Mod 10    'Earth Tones
           Next l
  Case 2:  For l = 1 To NumCountries
             MyMap.CountryColor(l) = 10    'White
           Next l
End Select

'Clear the previous display.
Call DrawBkg

'Draw the map!
MyMap.DisplayMap LandMap.hDC, Land.hDC, CFGBorders

End Sub

Private Sub StopBut_Click()

End Sub

Public Sub CollectMenuSettings()

'This subroutine grabs the checked info from the menus.

For i = 1 To 5
  If MenuSize(i).Checked = True Then CFGCountrySize = i
  If MenuPct(i).Checked = True Then CFGCheckedPct = i
  If MenuBorders(i).Checked = True Then CFGBorders = i
  If i <= 4 Then
    If MenuLakeSize(i).Checked = True Then CFGLakeSize = i
  End If
  If i <= 3 Then
    If MenuIslands(i).Checked = True Then CFGIslands = i
    If MenuShape(i).Checked = True Then CFGShape = i
    If MenuProp(i).Checked = True Then CFGProportion = i
  End If
  If i <= 2 Then
    If MenuColor(i).Checked = True Then CFGColor = i
    End If
Next i

'Now parse the menu settings.

'Get the number of countries that will fit in the
'selected area based on country size.
MaxCountrySize = Int((CFGCountrySize * 32.5) - 12.5)      '20-150 blocks per country.
Select Case CFGCheckedPct:
  Case 1:    LandPct = 0.1
  Case 2:    LandPct = 0.25
  Case 3:    LandPct = 0.5
  Case 4:    LandPct = 0.75
  Case 5:    LandPct = 0.9
End Select
NumCountries = Int((MyMap.Xsize * MyMap.Ysize * LandPct) / MaxCountrySize)
If NumCountries = 0 Then NumCountries = 1
If NumCountries > 998 Then NumCountries = 998

'Get the proportional size variance.
Select Case CFGProportion:
  Case 1:    PropPct = 0.1
  Case 2:    PropPct = 0.45
  Case 3:    PropPct = 0.75
End Select

'Get the minimum allowable lake size.
Select Case CFGLakeSize
  Case 1:    MinLakeSize = 0
  Case 2:    MinLakeSize = 5
  Case 3:    MinLakeSize = 10
  Case 4:    MinLakeSize = 20
End Select

'Get the percentage irregularity.
Select Case CFGShape
  Case 1:    ShapePct = 1
  Case 2:    ShapePct = 0.7
  Case 3:    ShapePct = 0.3
End Select

'Get the island parameters.
If CFGIslands = 2 Then
  CoastPctKeep = 0.99
  IslePctKeep = 0.01
End If


End Sub

Public Sub DrawBkg()

If CFGColor = 1 Then
  BitBlt hDC, 10, 0, 656, 472, Ocean.hDC, 0, 0, SRCCOPY
Else
  Land.Line (10 * Screen.TwipsPerPixelX, 0)-(657 * Screen.TwipsPerPixelX, 471 * Screen.TwipsPerPixelY), vbBlack, BF
End If
End Sub

Private Sub BorderBut_Click()

Dim Selected As Integer
Dim l As Integer
Dim m As Integer

Selected = Val(Selection.Text)

If Selected < 1 Or Selected > MyMap.LakeCode - 1 Then Selection.Text = "": Exit Sub

'If land is selected, find all non-water neighbors.
If Selected < 999 Then
  For l = 1 To MyMap.MaxNeighbors
    If MyMap.Neighbors(Selected, l) <> 0 And MyMap.Neighbors(Selected, l) < 999 Then
      MyMap.CountryColor(MyMap.Neighbors(Selected, l)) = 11
    End If
  Next l
End If

'If water is selected, find all countries who have this lake as a neighbor.
If Selected > 999 Then
  For l = 1 To NumCountries
    For m = 1 To MyMap.MaxNeighbors
      If MyMap.Neighbors(l, m) = Selected Then
        MyMap.CountryColor(l) = 11
      End If
    Next m
  Next l
End If
  
MyMap.DisplayMap LandMap.hDC, Land.hDC, CFGBorders

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim MapX As Integer
Dim MapY As Integer

If RollOver = False Or Selection.Enabled = False Then Exit Sub

  MapX = Int((Int(x / Screen.TwipsPerPixelX) - 10) / 8) + 1
  MapY = Int((Int(y / Screen.TwipsPerPixelY)) / 8) + 1
  
  If (MapX > 0) And (MapX <= MyMap.Xsize) And (MapY > 0) And (MapY <= MyMap.Ysize) Then
    If Button = 1 Then
      Selection.Text = CurrentMouse
      BorderBut.Enabled = True
    End If
  End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim MapX As Integer
Dim MapY As Integer

If RollOver = True Then
  'These instructions calculate the map coordinates where the mouse is.
  MapX = Int((Int(x / Screen.TwipsPerPixelX) - 10) / 8) + 1
  MapY = Int((Int(y / Screen.TwipsPerPixelY)) / 8) + 1
  
  If (MapX > 0) And (MapX <= MyMap.Xsize) And (MapY > 0) And (MapY <= MyMap.Ysize) Then
    CurrentMouse = MyMap.Grid(MapX, MapY)
    If CurrentMouse = 0 Or CurrentMouse = 999 Then
      Commentary.Caption = "Water"
    ElseIf CurrentMouse < 999 Then
      Commentary.Caption = "Country number " & CurrentMouse & " of " & NumCountries
    ElseIf CurrentMouse > 999 Then
      Commentary.Caption = "Water mass number " & CurrentMouse
    End If
  End If

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Unload Me
End

End Sub
