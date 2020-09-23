VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Orange Player"
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4965
   DrawWidth       =   9
   FillColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   118.285
   ScaleMode       =   0  'User
   ScaleWidth      =   106.969
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdReset 
      BackColor       =   &H80000009&
      Caption         =   "Reset"
      Height          =   495
      Left            =   3913
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   0
      Top             =   4560
      Width           =   1068
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
      End
      Begin VB.Menu oneplayer 
         Caption         =   "Players"
         Begin VB.Menu mnuoneplayer 
            Caption         =   "One Player"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnutwoplayers 
            Caption         =   "Two Players"
         End
      End
      Begin VB.Menu Mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu MnuInstructions 
         Caption         =   "Instructions"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'for storing point values
Private Type Coord
Enabled As Boolean
X As Integer
Y As Integer
R As Integer
G As Integer
B As Integer
End Type

'for storing the artificial intellegences points
Private Type tArtiesList
X As Integer
Y As Integer
End Type

'The squares(surrounded by 4 points and lines)
Private Type square
eSide(4) As Integer
Counted As Boolean
End Type


Dim TwoPlayers As Boolean 'whether the two player setting is on
Dim CP As Integer 'the current player
Dim ArtiesList() As tArtiesList 'the AI's list of possible squares
Dim T As Coord 'for mouseover
Dim OneDot As Coord 'for clicking
Dim M() As square 'the main array of squares
Dim Score(2) As Integer 'the scores
Dim ArtiesX, ArtiesY, ArtiesSide As Integer 'some values for the AI

'draw the graphics and set some values
Private Sub Form_Load()
Randomize
DrawBoard
DrawScore
TwoPlayers = False
CP = 1
End Sub

'reset the game(basically reload everything)
Private Sub CmdReset_Click()
If MsgBox("Are you sure you want to reset your game?", vbYesNo, "Reset?") = vbYes Then
    Score(0) = 0
    Score(1) = 0
    DrawBoard
    DrawScore
End If
End Sub

'wipe the form clean then draw the dots and reset the main array
Private Sub DrawBoard()
Me.Cls
For X = 0 To 10
    For Y = 0 To 10
        Me.PSet (3 + X * 10, 3 + Y * 10), RGB(255 - 3 * (X + Y), 0, 0)
    Next
Next
ReDim M(9, 9)
End Sub

'draw the fancy fading boxes and then print in the score for each player
Private Sub DrawScore()
Me.DrawWidth = 4
For i = 0 To 42 Step 0.5
    For j = 0 To 12 Step 0.5
        Me.PSet (42 + i, 106.702 + j), RGB(255, 255 - (i + 2 * j), 0)
    Next
Next
Me.CurrentX = 43.5
Me.CurrentY = 107
Me.Print "Orange Player: " & Score(1)
For i = 0 To 42 Step 0.5
    For j = 0 To 12 Step 0.5
        Me.PSet (i, 106.702 + j), RGB(0, 255 - 0.7 * (i + 2 * j), 0)
    Next
Next
Me.CurrentX = 1
Me.CurrentY = 107
If TwoPlayers Then
    Me.Print "Green Player: " & Score(0)
Else
    Me.Print "Computer: " & Score(0)
End If
Me.DrawWidth = 10
End Sub

'Put a new line into the main array
Private Function SetLine(X As Single, Y As Single, eX As Integer, eY As Integer, P As Integer)
On Error Resume Next
If X = eX Then
    If Y < eY Then
        M(X, Y).eSide(1) = P
        M(X - 1, Y).eSide(3) = P
    Else
        M(X, eY).eSide(1) = P
        M(X - 1, eY).eSide(3) = P
    End If
Else
    If X < eX Then
        M(X, Y).eSide(0) = P
        M(X, Y - 1).eSide(2) = P
    Else
        M(eX, Y).eSide(0) = P
        M(eX, Y - 1).eSide(2) = P
    End If
End If
End Function

'check whether there is a line where someone wants to draw one
Private Function GetLine(X As Single, Y As Single, eX As Integer, eY As Integer) As Boolean
On Error Resume Next
If X = eX Then
    If Y < eY Then
        If M(X, Y).eSide(1) = 0 Then GetLine = True: Exit Function
    Else
        If M(X, eY).eSide(1) = 0 Then GetLine = True: Exit Function
    End If
Else
    If X < eX Then
        If M(X, Y).eSide(0) = 0 Then GetLine = True: Exit Function
    Else
        If M(eX, Y).eSide(0) = 0 Then GetLine = True: Exit Function
    End If
End If
GetLine = False
End Function

'scan through the main array to see if there are any enclosed squares, then pass the play to the right person
Private Sub CheckSquares(Player As Integer)
Dim Z As Integer
Dim i, j As Integer
For i = 0 To 9
    For j = 0 To 9
        If M(i, j).eSide(0) <> 0 And M(i, j).eSide(1) <> 0 And M(i, j).eSide(2) <> 0 And M(i, j).eSide(3) <> 0 And M(i, j).Counted = False Then
            Call fillsquare(Player, i, j)
            Score(Player) = Score(Player) + 1
            M(i, j).Counted = True
            DrawScore
            currentplayer = Not currentplayer
            Z = Z + 1
        End If
    Next
Next
If TwoPlayers = True Then
    If Player = 1 And Z = 0 Then CP = 0: Me.Caption = "Green player"
    If Player = 1 And Z = 1 Then CP = 1: Me.Caption = "Orange player"
    If Player = 0 And Z = 0 Then CP = 1: Me.Caption = "Orange player"
    If Player = 0 And Z = 1 Then CP = 0: Me.Caption = "Green player"
Else
    If Player = 1 And Z = 0 Then Arty
    If Player = 0 And Z > 0 Then Arty
End If
End Sub

'fill in the centre of a square when somone wins it
Private Sub fillsquare(Player As Integer, ByVal X As Integer, ByVal Y As Integer)
If Player = 1 Then R = 255
If Player = 0 Then G = 255
If Player = 1 Then G = 200
Me.DrawWidth = 5
For i = 0 To 5 Step 0.5
    For j = 0 To 5 Step 0.5
        Me.PSet (i + 5.5 + X * 10, j * 0.82 + 6 + Y * 10), RGB(R, G - 5 * (i + j), 0)
    Next
Next
Me.DrawWidth = 10
End Sub

'draw a nice faded line
Private Sub drawline(X As Integer, Y As Integer, eX As Integer, eY As Integer, R As Integer, G As Integer)
Me.DrawWidth = 7
If X = eX Then
    If Y < eY Then
        For i = 0 To 10 Step 0.5
            Me.PSet (X, Y + i), RGB(R, G - 10 * i, 0)
        Next
    Else
        For i = 0 To 10 Step 0.5
            Me.PSet (X, Y - i), RGB(R, G - 10 * i, 0)
        Next
    End If
Else
    If X < eX Then
        For i = 0 To 10 Step 0.5
            Me.PSet (X + i, Y), RGB(R, G - 10 * i, 0)
        Next
    Else
        For i = 0 To 10 Step 0.5
            Me.PSet (X - i, Y), RGB(R, G - 10 * i, 0)
        Next
    End If
End If
Me.DrawWidth = 10
End Sub

'look where they have clicked and whether you need a line there
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.Point(X, Y) = RGB(128, 255, 0) Then
    X = X \ 10
    Y = Y \ 10
    If OneDot.Enabled = True Then
        If GetLine(X, Y, OneDot.X, OneDot.Y) Then GoTo Liner
    Else
        OneDot.R = 255
        OneDot.G = 255
        GoTo Norm
    End If
End If
If Me.Point(X, Y) = RGB(255, 168, 0) Then
    X = X \ 10
    Y = Y \ 10
    If OneDot.Enabled Then
Liner:
        If (OneDot.X = X And (OneDot.Y = Y + 1 Or OneDot.Y = Y - 1)) Or (OneDot.Y = Y And (OneDot.X = X + 1 Or OneDot.X = X - 1)) Then
            If CP = 1 Then
                Call drawline(3 + (X * 10), 3 + (Y * 10), 3 + (OneDot.X * 10), 3 + (OneDot.Y * 10), 255, 200)
            Else
                Call drawline(3 + (X * 10), 3 + (Y * 10), 3 + (OneDot.X * 10), 3 + (OneDot.Y * 10), 0, 255)
            End If
            Me.PSet (3 + (X * 10), 3 + (Y * 10)), RGB(255, 255, 0)
            Me.PSet (3 + (OneDot.X * 10), 3 + (OneDot.Y * 10)), RGB(255, 255, 0)
            Call SetLine(X, Y, OneDot.X, OneDot.Y, 2)
            T.Enabled = False
            OneDot.Enabled = False
            CheckSquares (CP)
        Else
            If OneDot.R = 255 Then
                Me.PSet (3 + (OneDot.X * 10), 3 + (OneDot.Y * 10)), RGB(255, 255, 0)
                OneDot.R = 0
                OneDot.G = 0
            Else
                Me.PSet (3 + (OneDot.X * 10), 3 + (OneDot.Y * 10)), RGB(255 - 3 * (X + Y), 0, 0)
            End If
            GoTo Norm
        End If
    Else
    
Norm:
        Me.PSet (3 + (X * 10), 3 + (Y * 10)), RGB(0, 0, 255)
        T.Enabled = False
        OneDot.Enabled = True
        OneDot.X = X
        OneDot.Y = Y
    End If
End If
End Sub

'make the rollover effect
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If T.Enabled Then
    If T.R = 255 Then
        Me.PSet (3 + (T.X * 10), 3 + (T.Y * 10)), RGB(255, 255, 0)
    Else
        Me.PSet (3 + (T.X * 10), 3 + (T.Y * 10)), RGB(255 - 3 * (T.X + T.Y), 0, 0)
    End If
End If
If Me.Point(X, Y) <= RGB(255, 0, 0) Or Me.Point(X, Y) = RGB(255, 255, 0) Then
    T.R = 0
    T.G = 0
    If Me.Point(X, Y) = RGB(255, 255, 0) Then
        T.R = 255
        T.G = 255
        X = X \ 10
        Y = Y \ 10
        'Me.Caption = X & ", " & Y  'useful for checking
        Me.PSet (3 + (X * 10), 3 + (Y * 10)), RGB(128, 255, 0)
    Else
        X = X \ 10
        Y = Y \ 10
        'Me.Caption = X & ", " & Y  'useful for checking
        Me.PSet (3 + (X * 10), 3 + (Y * 10)), RGB(255, 168, 0)
    End If
    T.X = X
    T.Y = Y
    T.Enabled = True
End If
End Sub


Private Sub MnuAbout_Click()
frmAbout.Show
End Sub

Private Sub MnuInstructions_Click()
Form2.Show
End Sub

Private Sub mnuoneplayer_Click()
mnuoneplayer.Checked = True
mnutwoplayers.Checked = False
TwoPlayers = False
CP = 1
Call CmdReset_Click
End Sub

Private Sub mnuReset_Click()
Call CmdReset_Click
End Sub

Private Sub mnutwoplayers_Click()
mnuoneplayer.Checked = False
mnutwoplayers.Checked = True
TwoPlayers = True
CP = 0
Call CmdReset_Click
End Sub











'This half of the code is for the artificial intellegence or 'Arty'
Private Sub Arty()
Randomize 'make the computers moves random
Dim Z As Integer
If CheckSquaresForArty = True Then GoTo DrawIt 'check whether there are 3 lines around any squares
                                            '
PickSide:                                   '
MakeArtiesList 'make a list of squares      '
Z = Int(Rnd * UBound(ArtiesList)) 'pick a   ' randomsquare out of the list
Call PickLine(Z) 'pick a side of the square '
                                            '
DrawIt: '<<<---------------------------------
    M(ArtiesX, ArtiesY).eSide(ArtiesSide) = 1 'set the chosen side in place
    Call PairArtiesSides(ArtiesX, ArtiesY, ArtiesSide) 'set the equivalent line in the square next to it
    Call DrawArtiesLine(ArtiesX, ArtiesY, ArtiesSide) 'actually draw the line
    CheckSquares (0) 'check for any full squares (function from above)
End Sub

'loop through all the squares and see if there are any with just three
Private Function CheckSquaresForArty() As Boolean
Dim RunTot As Integer
Dim i, j As Integer
For i = 0 To 9
    For j = 0 To 9
        For A = 0 To 3
            If M(i, j).eSide(A) <> 0 Then RunTot = RunTot + 1
        Next
        If RunTot = 3 Then
            For A = 0 To 3
                If M(i, j).eSide(A) = 0 Then ArtiesSide = A
            Next
            ArtiesX = i
            ArtiesY = j
            CheckSquaresForArty = True
            Exit Function
        Else
            RunTot = 0
        End If
    Next
Next
CheckSquaresForArty = False
End Function

'work out the real points for the line and send them to the main function above
Private Sub DrawArtiesLine(ByVal X As Integer, ByVal Y As Integer, ByVal side As Integer)
Dim Addition As Integer

Select Case side
Case 0
    GoTo Horizontal
Case 1
    GoTo Vertical
Case 2
    Addition = 10
    GoTo Horizontal
Case 3
    Addition = 10
    GoTo Vertical
End Select

Horizontal:
    Call drawline((X * 10) + 3, (Y * 10) + 3 + Addition, (X * 10) + 10 + 3, (Y * 10) + 3 + Addition, 0, 255)
    Me.PSet ((X * 10) + 3, (Y * 10) + 3 + Addition), RGB(255, 255, 0)
    Me.PSet ((X * 10) + 10 + 3, (Y * 10) + 3 + Addition), RGB(255, 255, 0)
    Exit Sub
Vertical:
    Call drawline((X * 10) + 3 + Addition, (Y * 10) + 3, (X * 10) + Addition + 3, (Y * 10) + 3 + 10, 0, 255)
    Me.PSet ((X * 10) + 3 + Addition, (Y * 10) + 3), RGB(255, 255, 0)
    Me.PSet ((X * 10) + Addition + 3, (Y * 10) + 3 + 10), RGB(255, 255, 0)
End Sub

'find the corresponding line and record it
Private Sub PairArtiesSides(ByVal X As Integer, ByVal Y As Integer, ByVal side As Integer)
On Error GoTo err
Select Case side
Case 0
    M(ArtiesX, ArtiesY - 1).eSide(2) = 1
Case 1
    M(ArtiesX - 1, ArtiesY).eSide(3) = 1
Case 2
    M(ArtiesX, ArtiesY + 1).eSide(0) = 1
Case 3
    M(ArtiesX + 1, ArtiesY).eSide(1) = 1
End Select
err:
End Sub

'make a list of possible good squares for arty to choose from.  Good means that it wont make
'a possible winning spot for the human
Private Sub MakeArtiesList()
Dim Z As Integer
Dim RunTot As Integer
For i = 0 To 9
    For j = 0 To 9
        For A = 0 To 3
            If M(i, j).eSide(A) <> 0 Then RunTot = RunTot + 1
        Next
        If RunTot < 2 Then
            ReDim Preserve ArtiesList(Z)
            ArtiesList(Z).X = i
            ArtiesList(Z).Y = j
            Z = Z + 1
        End If
        RunTot = 0
    Next
Next
If Z = 0 Then Call MakeArtiesSecondList
End Sub

'find any spots where there is a gap (even bad ones)
Private Sub MakeArtiesSecondList()
Dim Z As Integer
Dim RunTot As Integer
For i = 0 To 9
    For j = 0 To 9
        For A = 0 To 3
            If M(i, j).eSide(A) <> 0 Then RunTot = RunTot + 1
        Next
        If RunTot < 4 Then
            ReDim Preserve ArtiesList(Z)
            ArtiesList(Z).X = i
            ArtiesList(Z).Y = j
            RunTot = 0
            Z = Z + 1
        End If
        RunTot = 0
    Next
Next
End Sub

'Pick a random line from the chosen square
Private Sub PickLine(ByVal Num As Integer)
Dim Z(3) As Integer
Dim i As Integer
For A = 0 To 3
    If M(ArtiesList(Num).X, ArtiesList(Num).Y).eSide(A) = 0 Then
        Z(i) = A
        i = i + 1
    End If
Next
i = Int(Rnd * i)
ArtiesX = ArtiesList(Num).X
ArtiesY = ArtiesList(Num).Y
ArtiesSide = Z(i)
End Sub
