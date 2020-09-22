VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form fmrAlphaExample 
   Caption         =   "AlphaBlending Hot Chicks"
   ClientHeight    =   4905
   ClientLeft      =   2820
   ClientTop       =   630
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   7065
   Begin VB.Frame framLevel 
      Caption         =   "AlphaLvl"
      Height          =   3855
      Left            =   6000
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      Begin ComctlLib.Slider sldrAlphaLevel 
         Height          =   3375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   5953
         _Version        =   327682
         Orientation     =   1
         Max             =   255
         TickFrequency   =   15
      End
      Begin VB.Label lblAlphaLevel 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2040
         TabIndex        =   11
         Top             =   360
         Width           =   45
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "AlphaLevel"
      Height          =   975
      Left            =   3000
      TabIndex        =   6
      Top             =   3840
      Width           =   3015
      Begin ComctlLib.ProgressBar meterAlphaLevel 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   255
      End
      Begin VB.Label lblCurrAlpha 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   45
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "AutoAlpha"
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   3840
      Width           =   3015
      Begin ComctlLib.Slider sldrInterval 
         Height          =   615
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         _Version        =   327682
         Min             =   1
         Max             =   100
         SelStart        =   1
         TickFrequency   =   100
         Value           =   1
      End
      Begin VB.CommandButton cmdAutoAlpha 
         Caption         =   "||"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6120
      Top             =   4200
   End
   Begin VB.PictureBox Picture3 
      Height          =   855
      Left            =   6000
      Picture         =   "fmrAlphaExample.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   0
      Picture         =   "fmrAlphaExample.frx":1218C
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   3855
      Left            =   3000
      Picture         =   "fmrAlphaExample.frx":219EE
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "fmrAlphaExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================='
'=  Alpha Blending Tutorial By:     ='
'=      Aaron DeRenard              ='
'=  MSU-Fall04-Ex                   ='
'=----------------------------------='
'=  Free-Use Advisory:              ='
'=      Give credit where credit is ='
'=      Due!  Turn  me in as hw     ='
'=      And I find out              ='
'=      I'll report your ass!       ='
'===================================='


Private Type AlphaOptions
  AlphaOption As Byte
  AlphaFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
    Dim AO As AlphaOptions, newAO As Long
    Dim AlphaIncrease As Boolean

Private Sub cmdAutoAlpha_Click()
If Timer1.Enabled = True Then
    Timer1.Enabled = False
    cmdAutoAlpha.Caption = ">"
    framLevel.Visible = True
    sldrAlphaLevel.Value = AO.SourceConstantAlpha
    lblCurrAlpha.Caption = AO.SourceConstantAlpha
    lblAlphaLevel.Caption = AO.SourceConstantAlpha
ElseIf Timer1.Enabled = False Then
    Timer1.Enabled = True
    cmdAutoAlpha.Caption = "||"
    framLevel.Visible = False
End If
End Sub


Private Sub Form_Load()
    'Set the graphics mode to persistent
    Picture1.AutoRedraw = True
    Picture2.AutoRedraw = True
    'API uses pixels
    Picture1.ScaleMode = vbPixels
    Picture2.ScaleMode = vbPixels
    'set the parameters
    With AO
        .AlphaOption = AC_SRC_OVER
        .AlphaFlags = 0
        .SourceConstantAlpha = 0
        .AlphaFormat = 0
    End With
    'copy the AlphaOptions-structure to a Long
    RtlMoveMemory newAO, AO, 4
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    AlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, newAO
        AlphaIncrease = True
    Me.Show
End Sub

Private Sub meterAlphaLevel_Change()
    AO.SourceConstantAlpha = CByte(meterAlphaLevel.Value)
    Call CreateAlpha
    lblCurrAlpha.Caption = meterAlphaLevel.Value
End Sub

Public Sub CreateAlpha()
    'copy the AlphaOptions-structure to a Long
    RtlMoveMemory newAO, AO, 4
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    Picture2.Picture = Picture3.Picture
    AlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, newAO
End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub sldrAlphaLevel_Change()
    AO.SourceConstantAlpha = CByte(sldrAlphaLevel.Value)
    lblAlphaLevel.Caption = sldrAlphaLevel.Value
    Call CreateAlpha
End Sub

Private Sub sldrAlphaLevel_Click()
    AO.SourceConstantAlpha = CByte(sldrAlphaLevel.Value)
    lblAlphaLevel.Caption = sldrAlphaLevel.Value
    Call CreateAlpha
End Sub

Private Sub sldrAlphaLevel_Scroll()
    AO.SourceConstantAlpha = CByte(sldrAlphaLevel.Value)
    lblAlphaLevel.Caption = sldrAlphaLevel.Value
    Call CreateAlpha
End Sub

Private Sub sldrInterval_Change()
    Timer1.Interval = sldrInterval.Value
End Sub

Private Sub sldrInterval_Click()
    Timer1.Interval = sldrInterval.Value
End Sub

Private Sub Timer1_Timer()
    'Check to see if need increase of alpha or decrease
    If AlphaIncrease = True Then
        meterAlphaLevel.Value = meterAlphaLevel.Value + 1
        'Scale value on lblCurrAlpha to only show approx. multiples of 10
        If ((meterAlphaLevel.Value + 5) Mod 10) = 0 Then
            lblCurrAlpha.Caption = meterAlphaLevel.Value
        ElseIf meterAlphaLevel.Value = 0 Then
            lblCurrAlpha.Caption = scrlalphpalevel.Value
        End If
        'Check to see if at maximum possible Alpha level, if so, change mode from increase to decrease
        If meterAlphaLevel.Value = 255 Then
            AlphaIncrease = False
        End If
    ElseIf AlphaIncrease = False Then
        meterAlphaLevel.Value = meterAlphaLevel.Value - 1
        If ((meterAlphaLevel.Value + 5) Mod 10) = 0 Then
            lblCurrAlpha.Caption = meterAlphaLevel.Value
        ElseIf meterAlphaLevel.Value = 0 Then
            lblCurrAlpha.Caption = meterAlphaLevel.Value
        End If
        'Check to see if at minimum possible Alpha level, if so, change mode from decrease to increase
        If meterAlphaLevel.Value = 0 Then
            AlphaIncrease = True
        End If
    End If
    'Set the Alpha level to the byte value of meterAlphaLevel
    AO.SourceConstantAlpha = CByte(meterAlphaLevel.Value)
    'Call Alpha effect wrapper routine
    Call CreateAlpha
    
End Sub
