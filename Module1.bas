Attribute VB_Name = "Module1"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


Public Type POINTAPI
    x As Long
    y As Long
End Type

Enum ProgLong
    mVB = 1
    mCPlus = 2
    mDelphi = 3
    mJava = 4
    mPhotoShop = 5
End Enum

Private Type T_RGB
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

Private Type Pallet
    Sig As String * 3
    Pallet(0 To 8, 0 To 4) As Long
End Type

Public Const TileSize As Integer = 32
Public m_Pallet As Pallet
Public TemplateData(0 To 9) As Long
Public T_RGB As T_RGB

Public TDialog As New CDialog
Public Selected_LngCol As Long, ZoomFactor As Integer, Use_WebSafe As Boolean

Public Sub ShowPallet(PalBox As PictureBox, ColorSrc As PictureBox)
Dim x As Integer, y As Integer
Dim h As Integer, w As Integer

    'This function shows draws on the pallet
    w = (PalBox.ScaleWidth \ TileSize)
    h = (PalBox.ScaleHeight \ TileSize)
    
    For x = 0 To w
        For y = 0 To h
            ColorSrc.BackColor = m_Pallet.Pallet(x, y)
            BitBlt PalBox.hdc, x * TileSize, y * TileSize, TileSize, TileSize, ColorSrc.hdc, 0, 0, vbSrcCopy
        Next
    Next
    
    PalBox.Refresh

End Sub

Sub InitDialog(mHwnd As Long)
    'Init Dialog
    With TDialog
        .hInst = App.hInstance
        .DlgHwnd = mHwnd
        .flags = 0
        .InitialDir = FixPath(App.Path) & "pallets"
    End With
End Sub

Function StrRev(k As String)
Dim x As Integer, sBuff As String
    'Build a string backwards eg StrRev("ben")'returns neb
    For x = Len(k) To 1 Step -1
        sBuff = sBuff & Mid(k, x, 1)
    Next
    
    x = 0
    StrRev = sBuff
    
End Function

Public Function InvertColour(C As Long) As Long
Dim r As Long, g As Long, b As Long: r = C
    'Invert a color
    If r < 0 Then r = -r
    If r > 16777216 Then b = r \ 16777216: r = r - (b * 16777216)
    If r > 65535 Then b = r \ 65536: r = r - (b * 65536)
    If r > 255 Then g = r \ 256: r = r - (g * 256)
    InvertColour = RGB(-(r - 255), -(g - 255), -(b - 255))
End Function

Public Function WebSafe(intVal As Variant) As Integer
' I did not write this part found on the net somewere can't remmber'
' just like to say thanks for whoever did.
    Select Case intVal
        Case 0, 51, 102, 153, 204, 255
            WebSafe = intVal
        Case Else
            If intVal <= 26 Then
                WebSafe = 0: Exit Function
            ElseIf intVal > 26 And intVal <= 76 Then
                WebSafe = 51: Exit Function
            ElseIf intVal > 76 And intVal <= 127 Then
                WebSafe = 102: Exit Function
            ElseIf intVal > 127 And intVal <= 178 Then
                WebSafe = 153: Exit Function
            ElseIf intVal > 178 And intVal <= 229 Then
                WebSafe = 204: Exit Function
            ElseIf intVal > 229 Then
                WebSafe = 255: Exit Function
            End If
    End Select
End Function

Public Sub Gradient(TheObject As Object, Redval As Integer, Greenval As Integer, Blueval As Integer, TopToBottom As Boolean)
' I did not write this part found on the net somewere can't remmber'
' just like to say thanks for whoever did.

    'TheObject can be any object that supports the Line method (like forms and pictures).
    'Redval, Greenval, and Blueval are the Red, Green, and Blue starting values from 0 to 255.
    'TopToBottom determines whether the gradient will draw down or up.
    Dim Step%, Reps%, FillTop%, FillLeft%, FillRight%, FillBottom%, HColor$
    'This will create 63 steps in the gradient. This looks smooth on 16-bit and 24-bit color.
    'You can change this, but be careful. You can do some strange-looking stuff with it...
    Step = (TheObject.Height / 512)
    'This tells it whether to start on the top or the bottom and adjusts variables accordingly.
    If TopToBottom = True Then FillTop = 0 Else FillTop = TheObject.Height - Step
    FillLeft = 0
    FillRight = TheObject.Width
    FillBottom = FillTop + Step / 4
    'If you changed the number of steps, change the number of reps to match it.
    'If you don't, the gradient will look all funny.
    For Reps = 0 To 63
        'This draws the colored bar.
        Dim NewR As Integer, NewG As Integer, NewB As Integer
        If Use_WebSafe Then
            TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(WebSafe(Redval), WebSafe(Greenval), WebSafe(Blueval)), BF
        Else
            TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(Redval, Greenval, Blueval), BF
        End If
        
        Redval = Redval - 4
        Greenval = Greenval - 4
        Blueval = Blueval - 4
        'This prevents the RGB values from becoming negative, which causes a runtime error.
        If Redval <= 0 Then Redval = 0
        If Greenval <= 0 Then Greenval = 0
        If Blueval <= 0 Then Blueval = 0
        'More top or bottom stuff; Moves to next bar.
        If TopToBottom = True Then FillTop = FillBottom Else FillTop = FillTop - Step
        FillBottom = FillTop + Step
    Next
End Sub

Public Function Dec2Web(hDecCol As Long) As String
Dim StrHex As String
    'Convert a long color to a HTML color
    StrHex = Hex(hDecCol)
    Do While Len(StrHex) < 6
        StrHex = "0" & StrHex
        DoEvents
    Loop
    Dec2Web = "#" & Right(StrHex, 2) & Mid(StrHex, 3, 2) & Left(StrHex, 2)
    StrHex = ""
    
End Function

Public Function DoHex(hDecCol As Long, ProgLan As ProgLong) As String
Dim StrHex As String
    StrHex = Hex(hDecCol)
    
    Do While Len(StrHex) < 6
        StrHex = "0" & StrHex
        DoEvents
    Loop
    
    Select Case ProgLan
        Case mDelphi
            DoHex = "$00" & StrHex
        Case mVB
            DoHex = "&H" & StrHex & "&"
        Case mCPlus
            DoHex = "0x00" & StrHex
        Case mJava
            DoHex = "0x" & StrRev(Hex(hDecCol))
        Case mPhotoShop
            DoHex = StrRev(StrHex)
    End Select
    
End Function

Public Sub LongToRgb()
Dim mCol(2) As Byte
    'Converts a long color to a RGB color
    If (Selected_LngCol And &H80000000) Then Selected_LngCol = GetSysColor(Selected_LngCol And &HFFFFFF)
    
    CopyMemory mCol(0), Selected_LngCol, 3
    
    If Use_WebSafe Then
        T_RGB.Red = WebSafe(mCol(0))
        T_RGB.Green = WebSafe(mCol(1))
        T_RGB.Blue = WebSafe(mCol(2))
    Else
        T_RGB.Red = mCol(0)
        T_RGB.Green = mCol(1)
        T_RGB.Blue = mCol(2)
    End If
        
    Erase mCol
End Sub

Function FixPath(lzpath As String)
    If Right(lzpath, 1) = "\" Then FixPath = lzpath Else FixPath = lzpath & "\"
End Function

Function FindFile(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then FindFile = False Else FindFile = True
End Function

Function GetFileExt(lzFileName As String) As String
Dim x As Integer, i_pos As Integer
    'Returns a file names ext eg GetFileExt("ben.txt")'returns TXT
    For x = 1 To Len(lzFileName)
        If Mid(lzFileName, x, 1) = "." Then i_pos = x
    Next
    
    x = 0
    If i_pos = 0 Then GetFileExt = "": Exit Function
    GetFileExt = UCase(Mid(lzFileName, i_pos + 1, Len(lzFileName)))
    
End Function

Function SaveFile(lzFile As String, FileData() As Long)
Dim iFile As Long
    iFile = FreeFile
    'Saves data to a given filename
    Open lzFile For Binary As #iFile
        Put #iFile, , FileData()
    Close #iFile
    
End Function

Public Function OpenFile(FileName As String) As Variant
Dim iFile As Long
Dim mByte() As Byte
    'Opens a given file name and returns it's data
    iFile = FreeFile
    Open FileName For Binary As #iFile
        ReDim Preserve mByte(0 To LOF(iFile))
        Get #iFile, , mByte
    Close #iFile
    
    OpenFile = mByte
    Erase mByte
    
End Function
