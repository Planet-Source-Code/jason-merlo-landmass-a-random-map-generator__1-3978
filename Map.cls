VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Const XDIMENSION = 81
Const YDIMENSION = 59
Const UP = 1
Const DOWN = 2
Const LEFTY = 3
Const RIGHTY = 4
Const MAXIMUM_NEIGHBORS = 15

Dim Map(XDIMENSION, YDIMENSION) As Long
'Map Legend:
'0   = Plain water during build process (empty square)
'x   = Country number x
'999 = Coastline during build process
'1000 and up = Bodies of water (determined after land placement)

Dim Tent(2, 150) As Long
'Tent(1,a) = x coordinate of point a
'Tent(2,a) = y coordinate of point a

Dim Seed(XDIMENSION, YDIMENSION) As Single
'The seed matrix contains the details of the random parts of the map.
'This is for display purposes only -- Just a tweak.

Dim Neighbor() As Long
'The Neighbor array will be populated with data about which
'countries border others:
'For example:  z = Neighbor(x,y)
'x  = Country in question
'y  = Index number
'z  = The country number of a neighboring country.  Unused indexes return 0.
'There will be no more than MAXIMUM_NEIGHBORS possible borders detected,
'but this should change if the resolution is increased.

Dim CountryClr() As Long

Dim Direction As Long
Dim LastDirection As Long
Dim TryNumber As Long
Dim OneAway As Long
Dim Country As Long

Dim Done As Boolean
Dim CountryDone As Boolean
Dim FoundLand As Boolean
Dim FilledIn As Boolean

Dim i As Long
Dim j As Long
Dim k As Long
Dim x As Long
Dim y As Long
Dim ii As Long
Dim jj As Long
Dim xx As Long
Dim yy As Long

Public LakeCode As Long
Public NumCountries As Integer
Public MaxCountrySize As Integer
Public MinLakeSize As Integer

Dim CountrySize As Integer
Dim TentNum As Integer
Dim Block As Integer
Dim BlockTry As Integer
Dim FudgeCounter As Integer

Public Property Get Xsize() As Integer
    Xsize = XDIMENSION
End Property

Public Property Get Ysize() As Integer
    Ysize = YDIMENSION
End Property

Public Property Get MaxNeighbors() As Integer
    MaxNeighbors = MAXIMUM_NEIGHBORS
End Property

Public Property Get Grid(x As Integer, y As Integer) As Long
    Grid = Map(x, y)
End Property

Public Property Get Neighbors(x As Integer, y As Integer) As Long
    Neighbors = Neighbor(x, y)
End Property

Public Property Get CountryColor(x As Integer) As Integer
    CountryColor = CountryClr(x)
End Property

Public Property Let CountryColor(x As Integer, y As Integer)
    CountryClr(x) = y
End Property


Public Sub CreateMap(CountryCount As Integer, _
                    MaximumCountrySize As Integer, _
                    MinimumLakeSize As Integer, _
                    LandPct As Double, _
                    PropPct As Double, _
                    ShapePct As Double, _
                    CoastPctKeep As Double, _
                    IslePctKeep As Double, _
                    CFGColor As Integer, _
                    CFGIslands As Integer)

'We want some passed variables to stay public.
NumCountries = CountryCount
MaxCountrySize = MaximumCountrySize
MinLakeSize = MinimumLakeSize

'Redimension our arrays to save some memory.
ReDim Neighbor(NumCountries, MAXIMUM_NEIGHBORS)
ReDim CountryClr(NumCountries)

'All of map is considered coastline for placing of first country.
'Also initializing seed matrix.
For i = 1 To XDIMENSION
  For j = 1 To YDIMENSION
    Map(i, j) = 999
    Seed(i, j) = Rnd(1)
  Next j
Next i

Country = 0

Do While Country < NumCountries
Country = Country + 1

FudgeCounter = 0

'Clear out the last country by clearing out the tentative array.
For i = 1 To 2
  For j = 1 To MaxCountrySize
    Tent(i, j) = 0
  Next j
Next i

'Main country loop.
CountryDone = False
Do Until CountryDone = True

  'Find the starting position of the next country.
  Done = False
  Do Until Done = True
    x = Int(Rnd(1) * XDIMENSION) + 1
    y = Int(Rnd(1) * YDIMENSION) + 1
    'Take care of the None or Lots of Islands options.
    If ((CFGIslands = 1) And (Map(x, y) = 999)) Then
      Done = True
    End If
    If (CFGIslands = 3) And ((Map(x, y) = 0) Or (Map(x, y) = 999)) Then
      Done = True
    End If
    'The 'Some Islands' option is more complex.
    'If we found coastline, there's a large chance of keeping it.
    'If we found empty, there's a small chance of keeping it.
    If CFGIslands = 2 Then
      If ((Map(x, y) = 999) And (Rnd(1) < CoastPctKeep)) Or ((Map(x, y) = 0) And (Rnd(1) < IslePctKeep)) Then
        Done = True
      End If
    End If
  Loop
  Tent(1, 1) = x
  Tent(2, 1) = y
  
  TentNum = 1
  BlockTry = 0
  
  'Get a random size for this country.
  CountrySize = MaxCountrySize - Int(Rnd(1) * (MaxCountrySize * PropPct))

  'Now we search for blocks contiguous to this one.
  Do Until TentNum = CountrySize
    'Regular or Irregular?
    If Rnd(1) < ShapePct Then
      Block = Int(Rnd(1) * TentNum) + 1
    Else
      Block = TentNum
    End If
    'Get the coordinates of a block in this country.
    x = Tent(1, Block)
    y = Tent(2, Block)
    'Pick a random direction for a contiguous block.
    Call NewDirection
    Done = False
    Do Until (TryNumber = 5 Or Done = True)
      Select Case Direction
        Case UP
          OneAway = y - 1
          If OneAway > 0 Then
            'Check and see what's up.
            If ClearDirection(x, OneAway) = True Then
              'Map is clear in that direction.
              y = y - 1
              Call WriteToTent
              Done = True
            Else
              'Blocked.  Try new direction.
              Call NextDirection
            End If
          Else
            'We went off the top.  Try new direction.
            Call NextDirection
          End If
        Case DOWN
          OneAway = y + 1
          If OneAway <= YDIMENSION Then
            'Check and see what's down.
            If ClearDirection(x, OneAway) = True Then
              'Map is clear in that direction.
              y = y + 1
              Call WriteToTent
              Done = True
            Else
              'Blocked.  Try new direction.
              Call NextDirection
            End If
          Else
            'We went off the bottom.  Try new direction.
            Call NextDirection
          End If
        Case LEFTY
          OneAway = x - 1
          If OneAway > 0 Then
            'Check and see what's up.
            If ClearDirection(OneAway, y) = True Then
              'Map is clear in that direction.
              x = x - 1
              Call WriteToTent
              Done = True
            Else
              'Blocked.  Try new direction.
              Call NextDirection
            End If
          Else
            'We went off the left.  Try new direction.
            Call NextDirection
          End If
        Case RIGHTY
          OneAway = x + 1
          If OneAway <= XDIMENSION Then
            'Check and see what's up.
            If ClearDirection(OneAway, y) = True Then
              'Map is clear in that direction.
              x = x + 1
              Call WriteToTent
              Done = True
            Else
              'Blocked.  Try new direction.
              Call NextDirection
            End If
          Else
            'We went off the right.  Try new direction.
            Call NextDirection
          End If
      End Select
    Loop
    If TryNumber = 5 Then
      'This block is boxed in.  Try the next block in the country.
      Block = Block + 1
      If Block > TentNum Then
        Block = 1
      End If
      'Have we tried all blocks?
      BlockTry = BlockTry + 1
      If BlockTry = TentNum Then
        'This Country cannot fit.  Need new starting location.
        TentNum = CountrySize   'Fudging out of loop.
        FudgeCounter = FudgeCounter + 1
        'Move on to next country if we can't place this one
        'within a reasonable number of tries.
        If (FudgeCounter > 2000) Then
          CountryDone = True
          Country = Country - 1
          NumCountries = NumCountries - 1
        End If
      End If
    End If
  Loop
Loop

'Yay!  Our country is sitting in the Tent array.
'Let's copy it over to the Map.
'If we blew it last time, then we don't need to do anything here.
If FudgeCounter <= 2000 Then
  
For i = 1 To CountrySize
  Map(Tent(1, i), Tent(2, i)) = Country
Next i

If CFGColor = 1 Then
  CountryClr(Country) = Country Mod 10   'Earth Tones
Else
  CountryClr(Country) = 10   'White
End If

'Let's outline all countries on the map with coastline.
'Because if there are no islands, we must build on coastline next time.
For i = 1 To XDIMENSION
  For j = 1 To YDIMENSION
    'If this was the first country, we erase all coastline.
    If Country = 1 And Map(i, j) = 999 Then
      Map(i, j) = 0
    End If
    'If coastline is already there, it stays there.
    If Map(i, j) <> 999 Then
      FoundLand = False
      'Land up?
      x = i
      y = j - 1
      If y > 0 Then
        If Map(x, y) > 0 And Map(x, y) < 999 Then FoundLand = True
      End If
      'Land down?
      x = i
      y = j + 1
      If y <= YDIMENSION Then
        If Map(x, y) > 0 And Map(x, y) < 999 Then FoundLand = True
      End If
      'Land left?
      x = i - 1
      y = j
      If x > 0 Then
        If Map(x, y) > 0 And Map(x, y) < 999 Then FoundLand = True
      End If
      'Land right?
      x = i + 1
      y = j
      If x <= XDIMENSION Then
        If Map(x, y) > 0 And Map(x, y) < 999 Then FoundLand = True
      End If
      'Place coastline.
      If FoundLand = True And Map(i, j) = 0 Then
        Map(i, j) = 999
      End If
    End If
  Next j
Next i

End If

Loop    'Main country loop.

'Fill in lakes according to lakesize parameter.
Call FillLakes

'Populate the Neighbors array.
Call FindNeighbors

End Sub

Private Function WriteToTent()

'Found a good contiguous block, lets record its X and Y.
TentNum = TentNum + 1
Tent(1, TentNum) = x
Tent(2, TentNum) = y

'Was this the last tentative block?  If so, we need to signal we're done.
If TentNum = CountrySize Then
  CountryDone = True
End If

End Function

Private Function ClearDirection(Xcheck, Ycheck) As Boolean

'This function checks the Map and Tent arrays to see if
'the suggested block is already used.
ClearDirection = True

'Is the suggested block part of another country?
If Map(Xcheck, Ycheck) > 0 And Map(Xcheck, Ycheck) < 999 Then
  ClearDirection = False
End If

'Is the suggested block part of the current country (or lake)?
For i = 1 To TentNum
  If Tent(1, i) = Xcheck And Tent(2, i) = Ycheck Then
    ClearDirection = False
  End If
Next i

End Function

Private Function GetAdjacentColor(Xcheck, Ycheck)

If Map(Xcheck, Ycheck) <> 0 And Map(Xcheck, Ycheck) <> 999 And Country = 0 Then
  Country = Map(Xcheck, Ycheck)
End If

End Function

Private Function NextDirection()

'The last direction didn't work, so we try the next until we've done them all.
Direction = Direction + 1
If Direction = 5 Then Direction = 1
TryNumber = TryNumber + 1

End Function

Private Function NewDirection()

'Record which way we went last time.
LastDirection = Direction
If LastDirection < 1 Or LastDirection > 4 Then
  LastDirection = RandomDir
End If
Direction = RandomDir
'This will be our first try for a new way to go.
TryNumber = 1

End Function

Public Function DisplayMap(Source As Long, Dest As Long, CFGBorders As Integer)

Dim PieceNum As Integer
Dim LookUp As Integer
Dim LookDown As Integer
Dim LookLeft As Integer
Dim LookRight As Integer

'Lets update the map for everyone to see.
For ii = 1 To XDIMENSION
  For jj = 1 To YDIMENSION
    If Map(ii, jj) > 0 And Map(ii, jj) < 999 Then
      'This routine will determine which 'piece' gets used for this block.
      'Initialize our counters.
      PieceNum = 0
      LookUp = -1
      LookDown = -1
      LookLeft = -1
      LookRight = -1
      'See if there's water or coastline to the right of us.
      xx = ii + 1
      yy = jj
      If xx <= XDIMENSION Then
        LookRight = Map(xx, yy)
        If LookRight = 0 Or LookRight > 998 Then
          PieceNum = PieceNum + 1   'Add a binary 1
        End If
      End If
      'See if there's water or coastline to the left of us.
      xx = ii - 1
      yy = jj
      If xx > 0 Then
        LookLeft = Map(xx, yy)
        If LookLeft = 0 Or LookLeft > 998 Then
          PieceNum = PieceNum + 2   'Add a binary 2
        End If
      End If
      'See if there's water or coastline below us.
      xx = ii
      yy = jj + 1
      If yy <= YDIMENSION Then
        LookDown = Map(xx, yy)
        If LookDown = 0 Or LookDown > 998 Then
          PieceNum = PieceNum + 4   'Add a binary 4
        End If
      End If
      'See if there's water or coastline above us.
      xx = ii
      yy = jj - 1
      If yy > 0 Then
        LookUp = Map(xx, yy)
        If LookUp = 0 Or LookUp > 998 Then
          PieceNum = PieceNum + 8   'Add a binary 8
        End If
      End If
      Call DrawBlock(Source, Dest, ii, jj, CountryClr(Map(ii, jj)), PieceNum + (19 * Int(Seed(ii, jj) * 4)))
      
      'Now we check for the edit pieces between countries to smooth out borders.
      'Note that we *only* write on the current square.  We don't butt into
      'someone else's country!
      
      'All this does is smooth out a corner of our square if a country is diagonal
      'to us.
      
      'Check upper left.
      If (LookUp = LookLeft) And (LookUp <> Map(ii, jj)) And (LookUp > 0) And (LookUp < 999) Then
        Call DrawBlock(Source, Dest, ii, jj, CountryClr(LookUp), 15 + (19 * Int(Seed(ii, jj) * 4)))
      End If
      'Check upper right.
      If (LookRight = LookUp) And (LookRight <> Map(ii, jj)) And (LookRight > 0) And (LookRight < 999) Then
        Call DrawBlock(Source, Dest, ii, jj, CountryClr(LookRight), 16 + (19 * Int(Seed(ii, jj) * 4)))
      End If
      'Check lower left.
      If (LookLeft = LookDown) And (LookLeft <> Map(ii, jj)) And (LookLeft > 0) And (LookLeft < 999) Then
        Call DrawBlock(Source, Dest, ii, jj, CountryClr(LookLeft), 17 + (19 * Int(Seed(ii, jj) * 4)))
      End If
      'Check lower right.
      If (LookDown = LookRight) And (LookDown <> Map(ii, jj)) And (LookDown > 0) And (LookDown < 999) Then
        Call DrawBlock(Source, Dest, ii, jj, CountryClr(LookDown), 18 + (19 * Int(Seed(ii, jj) * 4)))
      End If
      
      'We will also add the selected border if two adjacent squares are
      'different countries.  If no border selected, then move on.
      
      If CFGBorders < 5 Then
        'Check up.
        If (LookUp <> Map(ii, jj)) And (LookUp > 0) And (LookUp < 999) Then
          Call DrawBlock(Source, Dest, ii, jj, 12, 0 + (CFGBorders - 1) * 4)
        End If
        'Check right.
        If (LookRight <> Map(ii, jj)) And (LookRight > 0) And (LookRight < 999) Then
          Call DrawBlock(Source, Dest, ii, jj, 12, 2 + (CFGBorders - 1) * 4)
        End If
        'Check left.
        If (LookLeft <> Map(ii, jj)) And (LookLeft > 0) And (LookLeft < 999) Then
          Call DrawBlock(Source, Dest, ii, jj, 12, 3 + (CFGBorders - 1) * 4)
        End If
        'Check down.
        If (LookDown <> Map(ii, jj)) And (LookDown > 0) And (LookDown < 999) Then
          Call DrawBlock(Source, Dest, ii, jj, 12, 1 + (CFGBorders - 1) * 4)
        End If
      End If
    End If
  Next jj
Next ii

End Function

Private Function DrawBlock(Source As Long, Dest As Long, x As Long, y As Long, Color As Long, PieceNum As Integer)

Dim PieceX As Integer
Dim ColorY As Integer

PieceX = 8 * PieceNum
ColorY = 16 * Color

BitBlt Dest, ((x - 1) * 8) + 10, (y - 1) * 8, 8, 8, Source, PieceX, ColorY + 8, SRCAND
BitBlt Dest, ((x - 1) * 8) + 10, (y - 1) * 8, 8, 8, Source, PieceX, ColorY, SRCINVERT

End Function

Private Function RandomDir()

'This function generates a random integer, 1-4.
RandomDir = Int(Rnd(1) * 4) + 1

End Function

Private Function FillLakes()

'This function checks for lakes on the entire map.
'If a lake is under the minimum size, then we fill it in
'with the color of a random adjacent country.
'If a lake is >= the minimum size, we will fill it in with
'the code for the next body of water (1000 and up).
'This will speed up the fill-in procedure so that each body
'of water is only tested once.  Also, we will be able to
'distinguish between lakes when we are done!

LakeCode = 1000

If MinLakeSize = 0 Then
  'If we chose No Lake Correction, we need to plant seeds so that we
  'can still identify bodies of water.
  For j = 1 To XDIMENSION
    For k = 1 To YDIMENSION
      If Map(j, k) = 999 Then
        Map(j, k) = LakeCode
        Call LabelBodyOfWater(LakeCode)
        LakeCode = LakeCode + 1
      End If
    Next k
  Next j
  Exit Function
End If

For j = 1 To XDIMENSION
  For k = 1 To YDIMENSION
    'We always start our lake search on coastline.
    FilledIn = False
    If Map(j, k) = 999 Then
      
      'Clear out the last lake by clearing out the tentative array.
      For i = 1 To MinLakeSize
        Tent(1, i) = 0
        Tent(2, i) = 0
      Next i
      
      Tent(1, 1) = j
      Tent(2, 1) = k
      TentNum = 1
      Country = 0
      Block = 1

      'Now we search for blocks contiguous to this one.
      Do Until TentNum = MinLakeSize
        'Get the coordinates of a block in this country.
        x = Tent(1, Block)
        y = Tent(2, Block)
        'Pick a direction to look for a contiguous block.
        Direction = 1
        Done = False
        Do Until (Direction = 5 Or Done = True)
        Select Case Direction
          Case UP
            OneAway = y - 1
            If OneAway > 0 Then
              'Check and see what's up.
              If ClearDirection(x, OneAway) = True Then
                'Map is clear in that direction.
                y = y - 1
                Call WriteToTent
                Done = True
              Else
                'Blocked.  Try new direction.
                Call GetAdjacentColor(x, OneAway)
                Direction = Direction + 1
              End If
            Else
              'We went off the top.  Try new direction.
              Direction = Direction + 1
            End If
          Case DOWN
            OneAway = y + 1
            If OneAway <= YDIMENSION Then
              'Check and see what's down.
              If ClearDirection(x, OneAway) = True Then
                'Map is clear in that direction.
                y = y + 1
                Call WriteToTent
                Done = True
              Else
                'Blocked.  Try new direction.
                Call GetAdjacentColor(x, OneAway)
                Direction = Direction + 1
              End If
            Else
              'We went off the bottom.  Try new direction.
              Direction = Direction + 1
            End If
          Case LEFTY
            OneAway = x - 1
            If OneAway > 0 Then
              'Check and see what's up.
              If ClearDirection(OneAway, y) = True Then
                'Map is clear in that direction.
                x = x - 1
                Call WriteToTent
                Done = True
              Else
                'Blocked.  Try new direction.
                Call GetAdjacentColor(OneAway, y)
                Direction = Direction + 1
              End If
            Else
              'We went off the left.  Try new direction.
              Direction = Direction + 1
            End If
          Case RIGHTY
            OneAway = x + 1
            If OneAway <= XDIMENSION Then
              'Check and see what's up.
              If ClearDirection(OneAway, y) = True Then
                'Map is clear in that direction.
                x = x + 1
                Call WriteToTent
                Done = True
              Else
                'Blocked.  Try new direction.
                Call GetAdjacentColor(OneAway, y)
                Direction = Direction + 1
              End If
            Else
              'We went off the right.  Try new direction.
              Direction = Direction + 1
            End If
        End Select
        Loop
        If Direction = 5 Then
          'This block is boxed in.  Try the next block in the lake.
          Block = Block + 1
          If Block > TentNum Then
            'Minimum lake can't fit here -- need to fill it in.
            For i = 1 To TentNum
              Map(Tent(1, i), Tent(2, i)) = Country
            Next i
            TentNum = MinLakeSize   'Fudging out of loop.
            FilledIn = True
          End If
        End If
        If Done = True Then
          Block = 1
        End If
      Loop
      If FilledIn = False Then
        'We've reached the minimum lake size.  Now we need to identify this
        'entire body of water with the next lake code.  'Note that we overwrite
        'our coastline here -- we don't need it anymore.
        
        'Part one of this procedure is to fill in the lake part we've
        'found already.
        For i = 1 To TentNum
          Map(Tent(1, i), Tent(2, i)) = LakeCode
        Next i
        
        'Part two of this procedure is to keep filling in the lake until
        'there is no more to fill in.  We do this with quick multiple passes.
        Call LabelBodyOfWater(LakeCode)
        LakeCode = LakeCode + 1
        
      End If
    End If
  Next k
Next j
End Function
Private Sub LabelBodyOfWater(LakeCode As Long)

        Done = False
        Do While Done = False
          Done = True  'If nothing gets written this pass, this will stay.
          For ii = 1 To XDIMENSION
            For jj = 1 To YDIMENSION
              If Map(ii, jj) = LakeCode Then
                'Check for all adjacent squares and fill them in if they are
                'water or coastline.  Note that we do diagonals here!
                xx = ii
                yy = jj
                'Check up.
                If yy - 1 > 0 Then
                  If Map(xx, yy - 1) = 0 Or Map(xx, yy - 1) = 999 Then
                    Map(xx, yy - 1) = LakeCode
                    Done = False
                  End If
                  'Check upper left.
                  If xx - 1 > 0 Then
                    If Map(xx - 1, yy - 1) = 0 Or Map(xx - 1, yy - 1) = 999 Then
                      Map(xx - 1, yy - 1) = LakeCode
                      Done = False
                    End If
                  End If
                End If
                'Check down.
                If yy + 1 <= YDIMENSION Then
                  If Map(xx, yy + 1) = 0 Or Map(xx, yy + 1) = 999 Then
                    Map(xx, yy + 1) = LakeCode
                    Done = False
                  End If
                  'Check lower right.
                  If xx + 1 <= XDIMENSION Then
                    If Map(xx + 1, yy + 1) = 0 Or Map(xx + 1, yy + 1) = 999 Then
                      Map(xx + 1, yy + 1) = LakeCode
                      Done = False
                    End If
                  End If
                End If
                'Check left.
                If xx - 1 > 0 Then
                  If Map(xx - 1, yy) = 0 Or Map(xx - 1, yy) = 999 Then
                    Map(xx - 1, yy) = LakeCode
                    Done = False
                  End If
                  'Check lower left.
                  If yy + 1 <= YDIMENSION Then
                    If Map(xx - 1, yy + 1) = 0 Or Map(xx - 1, yy + 1) = 999 Then
                      Map(xx - 1, yy + 1) = LakeCode
                      Done = False
                    End If
                  End If
                End If
                'Check right.
                If xx + 1 <= XDIMENSION Then
                  If Map(xx + 1, yy) = 0 Or Map(xx + 1, yy) = 999 Then
                    Map(xx + 1, yy) = LakeCode
                    Done = False
                  End If
                  'Check upper right.
                  If yy - 1 > 0 Then
                    If Map(xx + 1, yy - 1) = 0 Or Map(xx + 1, yy - 1) = 999 Then
                      Map(xx + 1, yy - 1) = LakeCode
                      Done = False
                    End If
                  End If
                End If
              End If
            Next jj
          Next ii
        Loop

End Sub

Private Sub FindNeighbors()

'This routine will populate the Neighbors array.
For i = 1 To NumCountries
  For j = 1 To MAXIMUM_NEIGHBORS
    Neighbor(i, j) = 0
  Next j
Next i

For i = 1 To XDIMENSION
  For j = 1 To YDIMENSION
    x = i
    y = j
    'See what's up.
    If y - 1 > 0 Then Call AddNeighbor(Map(x, y), Map(x, y - 1))
    'See what's down.
    If y + 1 <= YDIMENSION Then Call AddNeighbor(Map(x, y), Map(x, y + 1))
    'See what's left.
    If x - 1 > 0 Then Call AddNeighbor(Map(x, y), Map(x - 1, y))
    'See what's right.
    If x + 1 <= XDIMENSION Then Call AddNeighbor(Map(x, y), Map(x + 1, y))
  Next j
Next i

End Sub

Private Sub AddNeighbor(ToCountry As Long, NewNeighbor As Long)

'This sub works with FindNeighbors to populate the Neighbors array.

If ToCountry > 998 Then Exit Sub  'This is water!
If ToCountry = NewNeighbor Then Exit Sub  'Can't be our own neighbor!

Dim LastOne As Integer

LastOne = -1

For k = 1 To MAXIMUM_NEIGHBORS
  If Neighbor(ToCountry, k) = NewNeighbor Then Exit Sub  'Already there.
  If LastOne = -1 And Neighbor(ToCountry, k) = 0 Then LastOne = k
Next k

If LastOne = -1 Then MsgBox ("There were too many bordering countries found."): Exit Sub

Neighbor(ToCountry, LastOne) = NewNeighbor

End Sub
