VERSION 5.00
Object = "{08216199-47EA-11D3-9479-00AA006C473C}#2.1#0"; "RMCONTROL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Texturing Stuff"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Reset Size"
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      ItemData        =   "FrmMain.frx":0000
      Left            =   9120
      List            =   "FrmMain.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   210
      Width           =   975
   End
   Begin MSComDlg.CommonDialog D 
      Left            =   5520
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set"
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   7240
      TabIndex        =   12
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Animations"
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   30
      Width           =   2175
      Begin VB.OptionButton Option5 
         Caption         =   "None"
         Height          =   255
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Earth"
         Height          =   255
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Z"
         Height          =   255
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Y"
         Height          =   255
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "X"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.HScrollBar vSize 
      Height          =   255
      Left            =   840
      Max             =   10
      Min             =   1
      TabIndex        =   4
      Top             =   360
      Value           =   2
      Width           =   1455
   End
   Begin VB.HScrollBar hSize 
      Height          =   255
      Left            =   840
      Max             =   10
      Min             =   1
      TabIndex        =   1
      Top             =   120
      Value           =   2
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2760
      Top             =   1800
   End
   Begin RMControl7.RMCanvas RM 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   11245
   End
   Begin VB.Label Label4 
      Caption         =   "Color:"
      Height          =   255
      Left            =   8640
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Texture:"
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   " Width:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   " Height:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////
'/// 3D Texturing v2 - Effects & Basics      ///
'/// by: Solid - SOLID4K TEAM                   ///
'/// Plz Vote at PSC for 3D Texturing v2!  ///
'///////////////////////////////////////////////////

' Frame:    Holds data for a 3D Object
' Mesh:     What an object looks like 3D or "all around"
' Texture: The "skin" of an object, usually a picture.

Dim FR_Ball As Direct3DRMFrame3             ' Frame for holding data of the sphere
Dim MS_Ball As Direct3DRMMeshBuilder3   ' Mesh object for the sphere (what it looks like)
Dim TX_Ball As Direct3DRMTexture3          ' The texture of our sphere

Dim FR_Moon As Direct3DRMFrame3         ' Used in an animation later in this code
Dim MS_Moon As Direct3DRMMeshBuilder3 ' Mesh for the moon

Const Sin5 = 8.715574E-02!                      ' 5 Degrees
Const Cos5 = 0.9961947!                         ' 5 Degrees
Public Sub DX_Init() ' Initialize our objects
With RM
    .StartWindowed ' Start our 3D Scene
    .Viewport.SetBack 500 ' How far we can see back without objects disappearing
    .SceneFrame.SetSceneBackgroundRGB 0.3, 0.3, 0.3 ' Background color
    Set FR_Ball = .D3DRM.CreateFrame(.SceneFrame)   ' Create a new frame for the sphere
    Set FR_Moon = .D3DRM.CreateFrame(FR_Ball)          ' Create a new frame for the moon
End With
End Sub
Public Sub DX_MakeObjects() ' Make objects and visualizes them for the user to see
Set MS_Ball = RM.D3DRM.CreateMeshBuilder() ' Create a mesh builder for our sphere
Set TX_Ball = RM.D3DRM.LoadTexture(Text1.Text) ' Set the texture we'll be using

MS_Ball.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing ' This is the 3D object
MS_Ball.ScaleMesh 2, 2, 2   ' Make the sphere 2 times the size of it was originally
MS_Ball.SetTexture TX_Ball ' Adds the texture to the invisible sphere
FR_Ball.AddVisual MS_Ball   ' Makes the sphere and all its glory visible to the user!
MS_Ball.SetColorRGB 1, 1, 1 ' Make the ball white
End Sub

Private Sub Command1_Click() ' Load a coustom texture
' Note: some textures will not work, they must be directx size
' The rest of this sub is pretty self explanitory
D.DialogTitle = "Load Texture"
D.Filter = "Bitmap Files [.bmp]|*.bmp|Jpg Files [.jpg]|*.jpg|Jpeg Files [.jpeg]|*.jpeg|ALL Files|*.*"
D.FileName = ""
D.ShowOpen
If D.FileName = "" Then Exit Sub
Text1.Text = D.FileName
RM.SetFocus
End Sub

Private Sub Command2_Click()
' This sub makes the white fields in the texture the color selected
' Kinda hard to explain...
' Note: Instead of being the high factor of a color (255), it is a max of 1
FR_Ball.DeleteVisual MS_Ball ' Delete the sphere from the scene
DX_Init
DX_MakeObjects
If Combo1.ListIndex = 0 Then
    MS_Ball.SetColorRGB 1, 0, 0
ElseIf Combo1.ListIndex = 1 Then
    MS_Ball.SetColorRGB 0, 1, 0
ElseIf Combo1.ListIndex = 2 Then
    MS_Ball.SetColorRGB 0, 0, 1
ElseIf Combo1.ListIndex = 3 Then
    MS_Ball.SetColorRGB 1, 1, 1
End If
RM.SetFocus
End Sub

Private Sub Command3_Click()
' Refresh all surfaces and redraw our scene
' Note: this sub is equivilent to form_load
FR_Ball.DeleteVisual MS_Ball ' Delete the sphere from the scene
DX_Init
DX_MakeObjects
RM.SetFocus
End Sub

Private Sub Form_Load()
Text1.Text = App.Path & "\water.bmp"
DX_Init
DX_MakeObjects
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub hSize_Change()
FR_Ball.SetPosition Nothing, 0, 0, 0 ' Restores the original position - Note: 0 is center, -0 is left, and +0 is right
FR_Ball.DeleteVisual MS_Ball            ' Delete our sphere from the scene
DX_Init
DX_MakeObjects
MS_Ball.ScaleMesh hSize.Value / 2, vSize.Value / 2, 2 ' Change the sizes, X, Y, and Z
RM.SetFocus
End Sub

Private Sub Option1_Click()
' Note: SetRotation creates an animation by spinning the object until it is set to spin 0
FR_Ball.SetRotation FR_Ball, 0, Sin5, 0, 0.05 ' Rotate X
End Sub

Private Sub Option2_Click()
FR_Ball.SetRotation FR_Ball, Sin5, 0, 0, 0.05 ' Rotate Y
End Sub

Private Sub Option3_Click()
FR_Ball.SetRotation FR_Ball, 0, 0, Sin5, 0.05 ' Rotate Z
End Sub

Private Sub Option4_Click()
' This sub creates the moon, and puts it in the sphere's frame and spins it & the earth
Set MS_Moon = RM.D3DRM.CreateMeshBuilder()
MS_Moon.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing

Text1.Text = App.Path & "\earth.bmp"
Command2_Click

FR_Moon.SetPosition FR_Ball, 2, 1.5, 0
MS_Moon.ScaleMesh 0.2, 0.2, 0.2
FR_Moon.AddVisual MS_Moon
FR_Moon.SetRotation FR_Moon, 0, Sin5, 0, 0.05
FR_Ball.SetRotation FR_Ball, 0, -Sin5, 0, 0.01
End Sub

Private Sub Option5_Click()
' Stops the spinning, kills the moon, and deletes the earth texture
If FR_Moon.GetVisualCount <> 0 Then ' Check if the moon is visible
If Option4.Value = False Then               ' If the moon visible option isn't on
    FR_Moon.DeleteVisual MS_Moon        ' Delete the moon
    FR_Ball.DeleteVisual MS_Ball              ' Delete the main sphere
    DX_Init
    DX_MakeObjects
    Text1.Text = App.Path & "\water.bmp" ' Change the texture to default
    Command2_Click                                      ' Click the "Set" button
End If
End If
FR_Ball.SetRotation FR_Ball, 0, 0, 0, 0     ' Stops the spinning, if even spinning
End Sub

Private Sub RM_KeyDown(keyCode As Integer, Shift As Integer)
If keyCode = vbKeyLeft Then          ' Left arrow key pressed
    FR_Ball.SetOrientation FR_Ball, Sin5, 0, Cos5, 0, 1, 0 ' Spin sphere left
ElseIf keyCode = vbKeyRight Then ' Right arrow key pressed
    FR_Ball.SetOrientation FR_Ball, -Sin5, 0, Cos5, 0, 1, 0 ' Spin sphere right
ElseIf keyCode = vbKeyUp Then     ' Up arrow key pressed
    FR_Ball.SetOrientation FR_Ball, 0, -Sin5, Cos5, 0, Cos5, 0 ' Rotate sphere up
ElseIf keyCode = vbKeyDown Then ' Down arrow key pressed
    FR_Ball.SetOrientation FR_Ball, 0, Sin5, Cos5, 0, Cos5, 0   ' Rotate sphere down
End If
End Sub

Private Sub Timer1_Timer()
RM.Update ' Keeps our scene nice and updated
' Note:  This program doesn't work without this timer to keep everything updated
End Sub

Private Sub vSize_Change()
FR_Ball.SetPosition Nothing, 0, 0, 0    ' Resets the position of our sphere to the center
FR_Ball.DeleteVisual MS_Ball                ' Kills our sphere
DX_Init
DX_MakeObjects
MS_Ball.ScaleMesh hSize.Value / 2, vSize.Value / 2, 2 ' Sizes sphere
RM.SetFocus
End Sub
