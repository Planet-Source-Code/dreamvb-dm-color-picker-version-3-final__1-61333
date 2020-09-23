VERSION 5.00
Begin VB.UserControl Line3D 
   BackStyle       =   0  'Transparent
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
   ScaleHeight     =   90
   ScaleWidth      =   2745
   ToolboxBitmap   =   "3DLine.ctx":0000
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   540
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   540
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "Line3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum TLine
    Horizontal = 1
    Vertical = 2
End Enum

Const m_def_LineStyle = 1
Dim mLineStyle As Integer

Function DoLine()
    Select Case mLineStyle
        Case 1 ' Horizontal line
            Line1(0).X1 = 0
            Line1(0).X2 = 540
            Line1(0).Y1 = 0
            Line1(0).Y2 = 0
            '
            Line1(1).X1 = 0
            Line1(1).X2 = 540
            Line1(1).Y1 = 15
            Line1(1).Y2 = 15
            UserControl.Width = 2925
            UserControl_Resize
        Case 2
            Line1(0).X1 = 0
            Line1(0).X2 = 0
            Line1(0).Y1 = 0
            Line1(0).Y2 = 360
            '
            Line1(1).X1 = 15
            Line1(1).X2 = 15
            Line1(1).Y1 = 0
            Line1(1).Y2 = 360
            UserControl.Height = 1650
            UserControl_Resize
    End Select
    
End Function

Public Property Get LineStyle() As TLine
    LineStyle = mLineStyle
    If LineStyle = Horizontal Then
        mLineStyle = 1
    End If
    If LineStyle = Vertical Then
        mLineStyle = 2
    End If
    
End Property

Public Property Let LineStyle(ByVal vNewValue As TLine)
    mLineStyle = vNewValue
    PropertyChanged "LineStyle"
    DoLine
End Property

Private Sub UserControl_Initialize()
    mLineStyle = m_def_LineStyle
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mLineStyle = PropBag.ReadProperty("LineStyle", m_def_LineStyle)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    If mLineStyle = 1 Then
        UserControl.Height = 90
        Line1(0).X2 = UserControl.Width
        Line1(1).X2 = UserControl.Width
    End If
    
    If mLineStyle = 2 Then
        UserControl.Width = 90
        Line1(0).Y2 = UserControl.Height
        Line1(1).Y2 = UserControl.Height
    End If
    
End Sub

Private Sub UserControl_Show()
    DoLine
    UserControl_Resize
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("LineStyle", mLineStyle, m_def_LineStyle)
End Sub
