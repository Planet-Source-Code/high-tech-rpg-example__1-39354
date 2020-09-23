VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1980
      Left            =   960
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4920
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   4920
      Top             =   5520
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3840
      Picture         =   "Form1.frx":9042
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   514
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   7770
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   303
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   415
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Private Player_Direction(1 To 4) As Boolean
Private PlayerX As Integer
Private PlayerY As Integer
Private AnimationX As Integer
Private Animationy As Integer
Private PTile As Integer
'ptile is the tile that the player is
'is standing on, 0 - 125

Private Tiles(0 To 129) As tileZ

Private Type tileZ
    Property As String
End Type

'(C)ode Created By David W. Allen
'Feel Free to modify and use this code
'Add me to credits ;-)
'http://customsoftware.cjb.net
'email: techx@mailwire.net
'aim: h1ght3ch

Private Sub Form_Load()
For I = 0 To 129
Tiles(I).Property = "W"
Next I
LoadLevel "alevel"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then Player_Direction(1) = True
If KeyCode = vbKeyDown Then Player_Direction(2) = True
If KeyCode = vbKeyLeft Then Player_Direction(3) = True
If KeyCode = vbKeyRight Then Player_Direction(4) = True
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then Player_Direction(1) = False
If KeyCode = vbKeyDown Then Player_Direction(2) = False
If KeyCode = vbKeyLeft Then Player_Direction(3) = False
If KeyCode = vbKeyRight Then Player_Direction(4) = False
End Sub

Private Sub Timer1_Timer()
Dim ToX As Integer
Dim ToY As Integer

'take information from tile array
'and bit blt results to screen
'picture2.hdc :-)

For I = 0 To 129
If Int(PlayerX / 32) = Int(ToX / 32) And Int(PlayerY / 32) = Int(ToY / 32) Then PTile = I

If Tiles(I).Property = "G" Then x = 0
If Tiles(I).Property = "W" Then x = 32

BitBlt Picture1.hdc, ToX, ToY, 32, 32, Picture2.hdc, x, y, vbSrcCopy

If ToX > 370 Then
ToX = 0
ToY = ToY + 32
Else
ToX = ToX + 32
End If
Next I
'player is moving up
If Player_Direction(1) = True Then
If Tile_Can_Be_Walked_On(PTile) = True Then
PlayerY = PlayerY - 1
End If
Animationy = 0 'sets which row to use to draw player
ElseIf Player_Direction(2) = True Then 'down
If Tile_Can_Be_Walked_On(PTile) = True Then
PlayerY = PlayerY + 1
End If
Animationy = 32 'same
ElseIf Player_Direction(3) = True Then 'left
If Tile_Can_Be_Walked_On(PTile) = True Then
PlayerX = PlayerX - 1
End If
Animationy = 64 'same
ElseIf Player_Direction(4) = True Then 'right
If Tile_Can_Be_Walked_On(PTile) = True Then
PlayerX = PlayerX + 1
End If
Animationy = 96 'same
End If
'================draw player===================
TransparentBlt Picture1.hdc, PlayerX, PlayerY, 32, 32, Picture3.hdc, AnimationX, Animationy, 32, 32, vbBlue
'player 'walking' effect, scrolls across x axis

If AnimationX > 32 Then
        AnimationX = 0
    Else
        AnimationX = AnimationX + 32
End If
Picture1.Refresh
End Sub


Sub LoadLevel(FileName)
'if the level doesnt exist
'this will error ;-)
Dim AddSlash As String
Dim ATile As String
Dim TileInd As Integer 'tile number
Dim TileStr As String
Dim I As Integer
TileInd = 0
If Not Right(App.Path, 1) = "\" Then AddSlash = "\"
Open App.Path & AddSlash & FileName & ".txt" For Input As #1
Do While Not EOF(1)
        Input #1, TileStr
        For I = 1 To Len(TileStr)
        Tiles(TileInd).Property = Mid(TileStr, I, 1)
        x = Mid(TileStr, I, 1)
        TileInd = TileInd + 1
        Next I
Loop
Close #1

End Sub

Function Tile_Can_Be_Walked_On(TileNumber)
If Tiles(TileNumber).Property = "G" Then
'the tile was just grass
Tile_Can_Be_Walked_On = True
Else
'anything else, dont allow walking on
Tile_Can_Be_Walked_On = False
End If
End Function
