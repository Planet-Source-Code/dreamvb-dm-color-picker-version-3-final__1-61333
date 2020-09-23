VERSION 5.00
Begin VB.Form frmnew 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create new Pallet"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5280
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   4140
      TabIndex        =   6
      Top             =   150
      Width           =   930
   End
   Begin Project1.Flat2 Flat21 
      Height          =   1995
      Left            =   60
      TabIndex        =   5
      Top             =   135
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   3519
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   4140
      TabIndex        =   4
      Top             =   675
      Width           =   930
   End
   Begin VB.PictureBox PicCol 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   6780
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   4890
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicPallet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   90
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   2
      ToolTipText     =   "Click to add color"
      Top             =   165
      Width           =   3840
      Begin VB.Shape ShpMove 
         BorderColor     =   &H00808080&
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4140
      TabIndex        =   1
      Top             =   1725
      Width           =   930
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4140
      TabIndex        =   0
      Top             =   1200
      Width           =   930
   End
End
Attribute VB_Name = "frmnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX As Integer, OldY As Integer, bCanAdd As Boolean
Dim TempX As Integer, TempY As Integer

Sub BlankPallet()
Dim b As Boolean
    'Not very inportant just draws a default pallet
    For X = 0 To 8
        For Y = 0 To 4
            b = Not b
            m_Pallet.Pallet(X, Y) = ColorL(b)
        Next
    Next
    
    ShowPallet frmnew.PicPallet, frmnew.PicCol
End Sub

Private Sub OpenPallet(lpFile As String)
Dim nFile As Long
    nFile = FreeFile
    'Open a pallet file and display it
    Erase m_Pallet.Pallet
    m_Pallet.Sig = ""
    
    Open lpFile For Binary As #nFile
        Get #nFile, , m_Pallet
    Close #nFile
    
    If m_Pallet.Sig <> "Pal" Then
        MsgBox "The Pallet can not be loaded.", vbExclamation, Form1.Caption
    Else
        ShowPallet frmnew.PicPallet, frmnew.PicCol
    End If
    
End Sub

Sub DoColor()
    TDialog.flags = 2
    TDialog.ShowColor
    If TDialog.CancelError Then Exit Sub
    m_Pallet.Pallet(OldX \ TileSize, OldY \ TileSize) = TDialog.Color
    
    If bCanAdd Then
        PicCol.BackColor = TDialog.Color
        BitBlt PicPallet.hdc, TempX, TempY, 32, 32, PicCol.hdc, 0, 0, vbSrcCopy
        PicPallet.Refresh
    End If
    
End Sub

Sub GetCoords(X As Single, Y As Single)
    OldX = (X \ TileSize) * TileSize
    OldY = (Y \ TileSize) * TileSize
    
    If OldX < 0 Then OldX = 0
    If OldY < 0 Then OldY = Y
    If OldX > (TileSize * 8) Then OldX = (TileSize * 8) - TileSize
    If OldY > (TileSize * 4) - TileSize Then OldY = (TileSize * 4) - TileSize
    
    If ShpMove.Visible = False Then ShpMove.Visible = True
    ShpMove.Left = OldX
    ShpMove.Top = OldY
    
End Sub

Private Sub cmdclose_Click()
    Unload frmnew
End Sub

Private Sub cmdLoad_Click()
    TDialog.DialogTitle = "Open Pallet"
    TDialog.Filter = "Pallet Files (*.pal)" + Chr$(0) + "*.pal" + Chr$(0)
    TDialog.InitialDir = FixPath(App.Path) & "pallets\"
    TDialog.ShowOpen
    If Not TDialog.CancelError Then Exit Sub
    
    If GetFileExt(TDialog.FileName) <> "PAL" Then
        MsgBox "This is not a vaild Pallet filename.", vbCritical, "Invaild File Format"
        Exit Sub
    End If
    OpenPallet TDialog.FileName
End Sub

Private Sub cmdNew_Click()
    If MsgBox("This will clear the pallet are you sure you want to do this?", vbQuestion Or vbYesNo) = vbNo Then Exit Sub
    Erase m_Pallet.Pallet
    PicPallet.Cls
End Sub

Private Sub cmdsave_Click()
Dim sFile As String
    TDialog.DialogTitle = "Save Pallet"
    TDialog.Filter = "Pallet Files (*.pal)" + Chr$(0) + "*.pal" + Chr$(0)
    TDialog.InitialDir = FixPath(App.Path) & "pallets\"
    TDialog.ShowSave
    If Not TDialog.CancelError Then Exit Sub
    sFile = TDialog.FileName
    If Not GetFileExt(sFile) = "PAL" Then sFile = sFile & ".pal"
    
    m_Pallet.Sig = "Pal"
    Open sFile For Binary As #1
        Put #1, , m_Pallet
    Close #1
    
    'Clean up
    Erase m_Pallet.Pallet
    m_Pallet.Sig = ""
    sFile = ""

End Sub

Function ColorL(mCol) As Long
    If mCol Then ColorL = vbWhite Else ColorL = vbBlack
End Function

Private Sub Form_Load()
    InitDialog frmnew.Hwnd
    BlankPallet
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShpMove.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmnew = Nothing
    
End Sub

Private Sub PicPallet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    bCanAdd = True
    TempX = (X \ TileSize) * TileSize
    TempY = (Y \ TileSize) * TileSize
    Call DoColor
End Sub

Private Sub PicPallet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetCoords X, Y
End Sub

Private Sub PicPallet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bCanAdd = False
End Sub
