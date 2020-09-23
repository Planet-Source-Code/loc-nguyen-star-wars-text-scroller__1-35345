VERSION 5.00
Begin VB.Form FrmAndrew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Muahaha"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   617
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstScroll 
      Height          =   450
      Left            =   3240
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtScroll 
      Height          =   4815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "FrmAndrew.frx":0000
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox TxtSpeed 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Text            =   "50"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scroll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.PictureBox PicDraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   2640
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   433
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin VB.PictureBox PicBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   1
      Left            =   240
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   585
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   8775
   End
End
Attribute VB_Name = "FrmAndrew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal dreamAKA As Long) As Boolean

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal DWROP As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal DWROP As Long) As Long
    Private Const SRCCOPY = &HCC0020

Private NewBuffer As Object

Private Sub Form_Initialize()
    Set NewBuffer = Me.Controls("PicBuffer")
End Sub

Private Function AddBuffer() As PictureBox
    Dim n As Integer
    n = NewBuffer.ubound + 1
    Load NewBuffer(n)
    Set AddBuffer = NewBuffer(n)
End Function

Private Sub Command1_Click()
    ScrollFont PicDraw, TxtScroll.Text, 1.2, "Impact", 30, RGB(234, 215, 106), CInt(TxtSpeed.Text)
End Sub

'ScrollFont requires the controls PicBuffer(0), PicBuffer(1), and LstScroll
Private Sub ScrollFont(Pic As PictureBox, TheText As String, Slope As Double, TheFont As String, TheSize As Long, TheColor As Long, TheStep As Single, Optional IsBold As Boolean = False)
    'Array bounds = number of lines
    Dim Num(0 To 100) As Double, nBlend(0 To 100) As Double
    Dim tx As Double, ty As Double, tw As Double, th As Double, ax As Double
    Dim y(0 To 100) As Double, ox(0 To 100) As Double, oy(0 To 100) As Double, ow(0 To 100) As Double, oh(0 To 100) As Double
    Dim BigLst As Integer
    'Reset all objects' properties
    Pic.Cls
    LstScroll.Clear
    'Clear all the buffers
    For k = PicBuffer.LBound To PicBuffer.ubound
        PicBuffer(k).Cls
    Next k
    'Set initial properties for all buffers
    With PicBuffer(0)
        .FontName = TheFont
        .FontSize = TheSize
        .ForeColor = TheColor
        .FontBold = IsBold
    End With
    'Make an array of lines
    TempStr = TheText
    If Right(TempStr, 2) <> vbCrLf Then TempStr = TempStr & vbCrLf
    Do Until InStr(1, TempStr, vbCrLf, vbTextCompare) = 0
        RetPos = InStr(1, TempStr, vbCrLf, vbTextCompare)
        TheLine = Left(TempStr, RetPos - 1)
        TempStr = Right(TempStr, Len(TempStr) - (Len(TheLine) + 2))
        LstScroll.AddItem TheLine
    Loop
    For k = 2 To 2 * LstScroll.ListCount + 1
        'Add necessary buffers
        If PicBuffer.ubound < k Then AddBuffer
    Next k
    'Find largest string
    For k = 1 To LstScroll.ListCount - 1
        If PicBuffer(0).TextWidth(LstScroll.List(k)) > PicBuffer(0).TextWidth(LstScroll.List(BigLst)) Then BigLst = k
    Next k
    BigLst = PicBuffer(0).TextWidth(LstScroll.List(BigLst))
    If BigLst < Pic.Width Then BigLst = Pic.Width
    'Size all buffers to the max width and height
    PicBuffer(0).Width = BigLst
    PicBuffer(0).Height = PicBuffer(0).TextHeight(" ")
    For k = 2 To 2 * LstScroll.ListCount + 1
        PicBuffer(k).Width = BigLst
        PicBuffer(k).Height = PicBuffer(0).TextHeight(" ")
    Next k
    'Slant all buffers accordingly
    For k = 0 To LstScroll.ListCount - 1
        With PicBuffer(0)
            .Cls
            .CurrentX = BigLst / 2 - .TextWidth(LstScroll.List(k)) / 2
            .CurrentY = 0
        End With
        PicBuffer(0).Print LstScroll.List(k)
        For j = 0 To PicBuffer(0).Height - 1
            'Prevent Int() rounding (causes squiggly lines)
            x = Round(j / Slope, 0)
            nWidth = PicBuffer(0).Width - 2 * x
            StretchBlt PicBuffer(2 * (k + 1)).hdc, x, PicBuffer(2 * (k + 1)).Height - j, nWidth, 1, PicBuffer(0).hdc, 0, PicBuffer(0).Height - j, PicBuffer(0).Width, 1, SRCCOPY
        Next j
        y(k) = Pic.Height - 1 + k
    Next k
    'Where is the center?
    ax = Pic.Width / 2 - PicBuffer(0).Width / 2
    'Increment a certain y
    Do Until Num(LstScroll.ListCount - 1) >= 218
        For k = 0 To LstScroll.ListCount - 1
            If Num(k) >= 218 Then GoTo SkipDis
            BufInd = 2 * (k + 1)
            If y(k) >= Pic.Height Then
                If k > 0 Then y(k) = y(k - 1) + PicBuffer(BufInd - 1).Height
                GoTo SkipDis
            End If
            dy = Pic.Height - 1 - y(k)
            dx = dy / Slope
            nWidth = PicBuffer(BufInd).Width - 2 * dx
            ny = nWidth * PicBuffer(BufInd).Height / PicBuffer(BufInd).Width
            'Round to nearest pixel
            tx = ax + dx
            ty = y(k)
            tw = nWidth
            th = ny
            BitBlt Pic.hdc, Round(ox(k), 0), Round(oy(k), 0), Round(ow(k), 0), Round(oh(k), 0), PicBuffer(1).hdc, 0, 0, SRCCOPY
            StretchBlt PicBuffer(BufInd + 1).hdc, 0, 0, Round(tw, 0), Round(th, 0), PicBuffer(BufInd).hdc, 0, 0, PicBuffer(BufInd).Width, PicBuffer(BufInd).Height, SRCCOPY
            'AlphaBlend!
            Num(k) = 255 - 255 * th / PicBuffer(BufInd).Height
            nBlend(k) = vbBlue - CLng(Num(k)) * (vbYellow + 1)
            AlphaBlend Pic.hdc, Round(tx, 0), Round(ty, 0), Round(tw, 0), Round(th, 0), PicBuffer(BufInd + 1).hdc, 0, 0, Round(tw, 0), Round(th, 0), nBlend(k)
            'Store old values
            ox(k) = tx
            oy(k) = ty
            ow(k) = tw
            oh(k) = th
            y(k) = y(k) - th * TheStep / 500
'Skip loop if no rendering needed (save processing power)
SkipDis:
        Next k
        Pic.Refresh
        DoEvents
    Loop
End Sub

Function Round(x As Variant, DP As Integer) As Double
    Round = Int((x * 10 ^ DP) + 0.5) / 10 ^ DP
End Function

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
