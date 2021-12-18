VERSION 5.00
Begin VB.UserControl Progress 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Progress.ctx":0000
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3720
      ScaleHeight     =   1095
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   960
      ScaleHeight     =   735
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   240
      X2              =   240
      Y1              =   360
      Y2              =   1920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      X1              =   2520
      X2              =   120
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   2880
      X2              =   2880
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   480
      X2              =   2640
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Enumerations
Public Enum OrientationConst
    Horizontal
    Vertical
End Enum
'Default Property Values:
Const m_def_Orientation = 0
Const m_def_Max = 100
Const m_def_Position = 0
'Property Variables:
Dim m_Orientation As OrientationConst
Dim m_Max As Long
Dim m_Position As Long
'Event Declarations:
Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Occurs when any part of a form or PictureBox control is moved, enlarged, or exposed."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."

Private Sub UserControl_Resize()
    With Line1
        .X1 = 10
        .X2 = UserControl.Width - 10
        .Y1 = 10
        .Y2 = 10
    End With
    With Line2
        .X1 = UserControl.Width - 10
        .X2 = UserControl.Width - 10
        .Y1 = 10
        .Y2 = UserControl.Height - 10
    End With
    With Line3
        .X1 = UserControl.Width - 10
        .X2 = 10
        .Y1 = UserControl.Height - 10
        .Y2 = UserControl.Height - 10
    End With
    With Line4
        .X1 = 10
        .X2 = 10
        .Y1 = 10
        .Y2 = UserControl.Height - 10
    End With
    With Picture1
        .Left = 30
        .Top = 30
        .Width = UserControl.Width - 40
        .Height = UserControl.Height - 40
    End With
    With Picture2
        .Width = Picture1.Width
        .Height = Picture1.Height
    End With
    RaiseEvent Resize
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture2,Picture2,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Picture2.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Picture2.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,BackColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    FillColor = Picture1.BackColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    Picture1.BackColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = Picture1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Picture1.Picture = New_Picture
    PropertyChanged "Picture"
End Property

Public Property Get Position() As Long
    Position = m_Position
End Property

Public Property Let Position(ByVal New_Position As Long)
    On Error GoTo Ende
    m_Position = New_Position
    Picture1.Cls
    If Orientation = Vertical Then
        Picture1.PaintPicture Picture2.Image, 0, 0, , , 0, ((m_Position / m_Max) * Picture1.Height)
    Else
        Picture1.PaintPicture Picture2.Image, ((m_Position / m_Max) * Picture1.Height), 0, , , 0, 0
    End If
    PropertyChanged "Position"
    RaiseEvent Change
Ende:
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Position = m_def_Position
    m_Max = m_def_Max
    m_Orientation = m_def_Orientation
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Picture2.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Picture1.BackColor = PropBag.ReadProperty("FillColor", &H8000000D)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_Position = PropBag.ReadProperty("Position", m_def_Position)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    Set Picture2.Picture = PropBag.ReadProperty("Hintergrundbild", Nothing)
End Sub

Private Sub UserControl_Show()
    Picture1.PaintPicture Picture2.Image, 0, 0

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Picture2.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FillColor", Picture1.BackColor, &H8000000D)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Position", m_Position, m_def_Position)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("Hintergrundbild", Picture2.Picture, Nothing)
End Sub

Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

Public Property Get Orientation() As OrientationConst
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As OrientationConst)
    m_Orientation = New_Orientation
    PropertyChanged "Orientation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture2,Picture2,-1,Picture
Public Property Get Hintergrundbild() As Picture
Attribute Hintergrundbild.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Hintergrundbild = Picture2.Picture
End Property

Public Property Set Hintergrundbild(ByVal New_Hintergrundbild As Picture)
    Set Picture2.Picture = New_Hintergrundbild
    PropertyChanged "Hintergrundbild"
End Property

