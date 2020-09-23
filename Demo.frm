VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDemo 
   Caption         =   "Kath-Rock Software - (Gradient Demo)"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   Icon            =   "Demo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   3795
      Top             =   4410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      Height          =   3870
      Left            =   15
      ScaleHeight     =   3810
      ScaleWidth      =   3510
      TabIndex        =   6
      Top             =   1035
      Width           =   3570
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3720
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   4
      Top             =   2955
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picTools 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   0
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   306
      TabIndex        =   0
      Top             =   0
      Width           =   4590
      Begin VB.CommandButton cmdAngle 
         Caption         =   "&Angle"
         Height          =   855
         Left            =   2985
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "Color&2"
         Height          =   855
         Index           =   1
         Left            =   1935
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "Color&1"
         Height          =   855
         Index           =   0
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   30
         Width           =   735
      End
      Begin VB.Label lblLogo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kath-Rock"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   5
         Top             =   660
         Width           =   855
      End
      Begin VB.Image imgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   255
         Picture         =   "Demo.frx":1042
         Top             =   90
         Width           =   540
      End
   End
   Begin VB.Image imgTemp 
      Height          =   495
      Left            =   3720
      Top             =   3720
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfAngle         As Single
Private mlColor1        As Long
Private mlColor2        As Long
Private mbFormLoaded    As Boolean
Private mGradient       As New clsGradient



Private Sub DrawButtons()

    With picTemp
        .Width = 32
        .Height = 32
        .BackColor = vb3DFace
        .Cls
        
        'Color1 Button
        picTemp.Line (1, 1)-(.ScaleWidth - 1, .ScaleHeight - 1), vb3DShadow, B
        picTemp.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), vb3DHighlight, B
        picTemp.Line (0, 0)-(.ScaleWidth, .ScaleHeight), vb3DDKShadow, B
        Set imgTemp.Picture = picTemp.Image
        Set cmdAngle.Picture = imgTemp.Picture
        
        'Color2 Button
        picTemp.Line (2, 2)-(.ScaleWidth - 2, .ScaleHeight - 2), mGradient.Color1, BF
        Set imgTemp.Picture = picTemp.Image
        Set cmdColor(0).Picture = imgTemp.Picture
        picTemp.Line (2, 2)-(.ScaleWidth - 2, .ScaleHeight - 2), mGradient.Color2, BF
        Set imgTemp.Picture = picTemp.Image
        Set cmdColor(1).Picture = imgTemp.Picture
        
        'Angle Button
        mGradient.Draw picTemp
        picTemp.Line (1, 1)-(.ScaleWidth - 1, .ScaleHeight - 1), vb3DShadow, B
        picTemp.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), vb3DHighlight, B
        picTemp.Line (0, 0)-(.ScaleWidth, .ScaleHeight), vb3DDKShadow, B
        Call DrawAngle(picTemp, mGradient.Angle)
        Set imgTemp.Picture = picTemp.Image
        Set cmdAngle.Picture = imgTemp.Picture
        
    End With
    
End Sub

Private Sub DrawGradient()

    With mGradient
        .Angle = mfAngle
        .Color1 = mlColor1
        .Color2 = mlColor2
        .Draw picDraw
    End With
    
    picDraw.Refresh
    
End Sub


Private Sub cmdAngle_Click()

    If frmCustomAngle.Display(Me, mfAngle, mlColor1, mlColor2) Then
        Call DrawGradient
        Call DrawButtons
    End If
    
End Sub

Private Sub cmdColor_Click(Index As Integer)

Dim lColor As Long

    If Index = 0 Then
        lColor = mlColor1
    Else
        lColor = mlColor2
    End If
    
    On Error GoTo LocalError
    
    With cdlColor
        .CancelError = True
        .Flags = cdlCCRGBInit
        .Color = lColor
        .ShowColor
        lColor = .Color
    End With
    
    If Index = 0 Then
        mlColor1 = lColor
    Else
        mlColor2 = lColor
    End If
    
    Call DrawGradient
    Call DrawButtons
    
NormalExit:
    Exit Sub
    
LocalError:
    If Err.Number <> cdlCancel Then
        MsgBox Err.Description, vbExclamation
    End If
    Resume NormalExit
    
End Sub


Private Sub Form_Load()

    With picTools
        .BackColor = vb3DFace
        .Cls
    End With
        
    With mGradient
        mfAngle = .Angle
        mlColor1 = .Color1
        mlColor2 = .Color2
    End With
    
    Call DrawButtons
    mbFormLoaded = True
    
End Sub


Private Sub Form_Resize()

Const lMinWidth     As Long = 4080
Const lMinHeight    As Long = 2790

    If Me.WindowState <> vbMinimized Then
        If Me.Width < lMinWidth Then
            Me.Width = lMinWidth
        ElseIf Me.Height < lMinHeight Then
            Me.Height = lMinHeight
        Else
            picDraw.Move 0, picTools.Height, Me.ScaleWidth, Me.ScaleHeight - picTools.Height
            Call DrawGradient
        End If
    End If
    
End Sub


