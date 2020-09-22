VERSION 5.00
Object = "{DF6D6558-5B0C-11D3-9396-008029E9B3A6}#1.0#0"; "ezVidC60.ocx"
Begin VB.Form MainFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Motion Detector by Sebastjan Dobnik"
   ClientHeight    =   8190
   ClientLeft      =   2850
   ClientTop       =   1440
   ClientWidth     =   9720
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   648
   ShowInTaskbar   =   0   'False
   Tag             =   "9o"
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      LargeChange     =   100
      Left            =   3180
      Max             =   2500
      Min             =   100
      TabIndex        =   6
      Top             =   7800
      Value           =   800
      Width           =   4755
   End
   Begin VB.PictureBox MotionCon 
      BackColor       =   &H00CE9A6F&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   60
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   4
      Top             =   60
      Width           =   9600
      Begin VB.PictureBox MotionPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   3600
         Left            =   2700
         ScaleHeight     =   238
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   318
         TabIndex        =   5
         Top             =   1920
         Width           =   4800
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "STARTED"
      Height          =   735
      Index           =   5
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7380
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Video Format"
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   7620
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Video Source"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   7620
      Width           =   1395
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   180
      Top             =   7140
   End
   Begin vbVidC60.ezVidCap VidCap 
      Height          =   540
      Left            =   9780
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   953
      AutoSize        =   0   'False
      BackColor       =   15299894
      BorderStyle     =   0
      StreamMaster    =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tolerance:"
      Height          =   195
      Left            =   3180
      TabIndex        =   8
      Top             =   7500
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   195
      Left            =   4020
      TabIndex        =   7
      Top             =   7500
      Width           =   480
   End
   Begin VB.Menu MenExit 
      Caption         =   "Exit"
      Visible         =   0   'False
      Begin VB.Menu MenExitExit 
         Caption         =   "Konec"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mdTriger As Single                  'motion triger level
Dim mdSample(50, 50, 250) As Single     'color map for pixel comparisson

Sub GetMotion()
    Dim ColorSumStr As String           'sum of pixel color
    Dim ColorRedStr As String           'red
    Dim ColorGreenStr As String         'green
    Dim ColorBlueStr As String          'blue
    Dim ColorRedDec As Single           'red
    Dim ColorGreenDec As Single         'green
    Dim ColorBlueDec As Single          'blue
    Dim PixX As Single                  'curent pixel X
    Dim PixY As Single                  'curent pixel Y
    Dim AveragePixel(5) As Single       'Average color from 6 pixels
    Static Counter As Single            'counter
    Dim AverageSum As Single            'Average sum of all colors
    
    Dim BoxesX As Single                'how many 'detection boxes - x axis
    Dim BoxesY As Single                'how many 'detection boxes - y axis
    Dim AveragePixelLoop As Single      'defines how many frames does this sub compare
    
    BoxesX = 16                         'from 1 to 50
    BoxesY = 16                         'from 1 to 50
    AveragePixelLoop = 30               'from 1 to 250
    
    Dim Repeat As Single
    Dim Px As Single, Py As Single
    
    For Px = 0 To (MotionPic.Width) Step Int(MotionPic.Width / BoxesX)
    For Py = 0 To (MotionPic.Height) Step Int(MotionPic.Height / BoxesY)
            
            PixX = Fix(Px / (MotionPic.Width / BoxesX))
            PixY = Fix(Py / (MotionPic.Height / BoxesY))
            For Repeat = 0 To 5
                ColorSumStr = Right$("000000" + Hex(GetPixel(MotionPic.hdc, Px + Repeat, Py + Repeat)), 6)
                ColorRedStr = Mid$(ColorSumStr, 5, 2)
                ColorGreenStr = Mid$(ColorSumStr, 3, 2)
                ColorBlueStr = Mid$(ColorSumStr, 1, 2)
                ColorRedDec = Val("&H" + ColorRedStr)
                ColorGreenDec = Val("&H" + ColorGreenStr)
                ColorBlueDec = Val("&H" + ColorBlueStr)
                AveragePixel(Repeat) = ColorRedDec + ColorGreenDec + ColorBlueDec
            Next
            
            Counter = Counter + 1
            If Counter = AveragePixelLoop Then Counter = 1
            
            mdSample(PixX, PixY, 0) = 0
            mdSample(PixX, PixY, Counter) = 0
            For Repeat = 0 To 5
                mdSample(PixX, PixY, 0) = mdSample(PixX, PixY, 0) + AveragePixel(Repeat)
                mdSample(PixX, PixY, Counter) = mdSample(PixX, PixY, 0) + AveragePixel(Repeat)
            Next
            
            AverageSum = 0
            For Repeat = 1 To AveragePixelLoop
                AverageSum = AverageSum + mdSample(PixX, PixY, Repeat)
            Next
            AverageSum = AverageSum / AveragePixelLoop
            
            
            'preveri proženje motion-a
            If Abs(mdSample(PixX, PixY, 0) - AverageSum) > mdTriger Then
                MotionPic.Line (Px - 4, Py - 4)-Step((MotionPic.Width / BoxesX) - 4, (MotionPic.Height / BoxesY) - 4), , B
            End If
    Next
    Next
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        VidCap.ShowDlgVideoSource
    ElseIf Index = 1 Then
        VidCap.ShowDlgVideoFormat
    ElseIf Index = 2 Then
        VidCap.StreamNoFile
    ElseIf Index = 5 Then
        Timer1.Enabled = Not Timer1.Enabled
        If Timer1.Enabled Then
            Command1(Index).Caption = "STARTED"
            Command1(Index).BackColor = &H80FF80
        Else
            Command1(Index).Caption = "STOPPED"
            Command1(Index).BackColor = &HFF&
            VidCap.Preview = True
        End If
    End If
End Sub


Private Sub Form_Load()
    'set triger lever to 700
    HScroll1.Value = 700
End Sub

Private Sub HScroll1_Change()
    'set triger lever
    mdTriger = HScroll1.Value
    Label4 = Format$(HScroll1.Value)
End Sub

Private Sub HScroll1_Scroll()
    'set triger lever
    mdTriger = HScroll1.Value
    Label4 = Format$(HScroll1.Value)
End Sub

Private Sub MotionPic_Resize()
    'resize and cente picture
    MotionPic.Left = (MotionCon.Width - MotionPic.Width) / 2
    MotionPic.TOp = (MotionCon.Height - MotionPic.Height) / 2
End Sub

Private Sub Timer1_Timer()
    'main loop
    If VidCap.CapSingleFrame Then
        MotionPic.Cls
        VidCap.SaveDIB VidCap.CaptureFile
        Set MotionPic.Picture = LoadPicture(VidCap.CaptureFile)
        GetMotion
        Kill VidCap.CaptureFile
    End If
End Sub

