VERSION 5.00
Begin VB.UserControl Trans 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1230
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   75
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   82
   ToolboxBitmap   =   "Obj Trans.ctx":0000
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4125
      Left            =   600
      ScaleHeight     =   275
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   4125
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   120
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Image Image2 
      Height          =   4125
      Left            =   5040
      Top             =   2520
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      FillColor       =   &H00808080&
      Height          =   975
      Left            =   70
      Top             =   65
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Object Transpancy Control- By Marco Samy"
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Trans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////
'////             Tansparent Object Control                 ////
'////             Code By: Marco Samy Nasif                 ////
'////             mail:marco_s2@hotmail.com                 ////
'////             Call:  (+20) 12 72 42 974                 ////
'////             /////////////////////////                 ////
'////             /////////////////////////                 ////
'////             Arabic Republic Of  EGYPT                 ////
'////             Copyright (c)2002,   FREE                 ////
'////             To  Use  Or Include  into                 ////
'////             Your    Own    Programs .                 ////
'///////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////

'4 Neede Api Function .. (1) Combine two regions, (2) Conver an Object Space to a fixed region , (3) to know a point color from an DC-Compitable Object, (4) Delete Unused Data to free resources , (5) Creat a rectangular region .. (c) By Marco Samy 2002
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal nXPos As Long, ByVal nYPos As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Const RGN_AND As Long = 1
Private Const RGN_COPY As Long = 5
Private Const RGN_DIFF As Long = 4
Private Const RGN_MAX As Long = 2
Private Const RGN_MIN As Long = RGN_AND
Private Const RGN_OR As Long = 2
Private Const RGN_XOR As Long = 3
Const m_def_CurrentRGN = 0
Const m_def_ScaleX = 1
Const m_def_ScaleY = 1
Const m_def_BitColor = 0
Dim m_CurrentRGN As Long
Dim m_ScaleX As Variant
Dim m_ScaleY As Variant
Dim m_MyObject As Object
Dim m_BitColor As OLE_COLOR
Event BitMapLoaded() 'This Event Raised when the bitmap converted to region is done
Private Sub UserControl_Resize()
UserControl.Height = Shape1.Height * Screen.TwipsPerPixelY
UserControl.Width = Shape1.Width * Screen.TwipsPerPixelX
End Sub
Public Property Get CurrentRGN() As Long
Attribute CurrentRGN.VB_MemberFlags = "400"
    CurrentRGN = m_CurrentRGN
End Property
Public Function DeleteRGN() As Variant
'delete rgn is necessary for memory
'this will free up memory
DeleteObject m_CurrentRGN 'the m_currentRGn is a value to handle a windows resource
m_CurrentRGN = 0 'setting no region created.
End Function
Public Function DoObject() As Variant
If m_CurrentRGN = 0 Then GoTo err2
On Error GoTo errp 'on error the object is has no hwnd, so we can customize a region to it, raise error
SetWindowRgn m_MyObject.hwnd, m_CurrentRGN, True 'applying the created region to fixed window
Exit Function
errp:
Err.Clear
Err.Description = "Selected Object Dosent Allow Transpancy, Call 0127242974."
Err.Raise 1533
Exit Function
err2:
Err.Clear
Err.Description = "There are no BitMap Region to apply, Call 0127242974."
Err.Raise 1533
End Function
Public Function DoToObject(sHandel As Long) As Variant
'we must have a region created
If m_CurrentRGN = 0 Then GoTo err2 'no region created, raise error message
SetWindowRgn sHandel, m_CurrentRGN, True 'applying created region to a fixed object
Exit Function
err2:
Err.Clear
Err.Description = "There are no BitMap Region to apply, Call 0127242974."
Err.Raise 1533
End Function
Public Function LoadBitmap(ByVal sFileName As String) As Long
'What we will do?
'this function to set the working region m_CurrentRGN
'the value of this variable is not the region but it's a value of handle of the region in the windows memory resources
'we can create alot of regions and attach it together to make the final region what we will
'aply to the form/object using SetWindowRgn [API Function]
'---
'how we can create a region from image?
'using a fixed transparent color
'by the following code we will scan every pixel on the given photo
'and when it match the transparent color we will put start point here
'the system will resume scannig, when it reaches non transparent color
'it will craete a region there and this region will be attached to a general
'region which well be our form region.
DeleteRGN 'checking deleting resoures, freeup memory
Image1.Stretch = False 'unstrectch
Image1.Picture = LoadPicture(sFileName)
DoEvents
P1.Height = Image1.Height
P1.Width = Image1.Width
P1.Picture = Image1.Picture
P1.Refresh
'declaring needed variables
Dim bX As Long
Dim cX As Long
Dim cY As Long
Dim MainRGN As Long
Dim FlyRegion As Long
Dim srcWidth
Dim srcHeught
srcWidth = P1.Width
SrcHeight = P1.Height
'here we start scannig
bX = -1
For cY = 0 To SrcHeight - 1 'scanning vertically
For cX = 0 To srcWidth - 1 'scanning horizontally
If GetPixel(P1.hDC, cX, cY) = m_BitColor Then 'getpixel is faster than .Point(X,Y)
If bX = -1 Then bX = cX
If cX = srcWidth - 1 Then
If MainRGN = 0 Then 'if no region was created yet(the general region)
'craete a region using provided data
MainRGN = CreateRectRgn(bX * m_ScaleX, cY * m_ScaleY, (cX + 1) * m_ScaleX, (cY + 1) * m_ScaleY)
Else 'there was region creaed
'so we will create a temporary region (FlyRegion)
'and we will attach it to the general region
'delete the temporary region (FlyRegion)
FlyRegion = CreateRectRgn(bX * m_ScaleX, cY * m_ScaleY, (cX + 1) * m_ScaleX, (cY + 1) * m_ScaleY)
CombineRgn MainRGN, MainRGN, FlyRegion, RGN_MAX 'attach the two regions together
DeleteObject FlyRegion 'deleting temporary
End If
bX = -1 'starting again
End If
Else
If bX <> -1 Then
If MainRGN = 0 Then 'no region created yet
'create the main region
MainRGN = CreateRectRgn(bX * m_ScaleX, cY * m_ScaleY, cX * m_ScaleX, (cY + 1) * m_ScaleY)
Else
'attach the two regions together
FlyRegion = CreateRectRgn((bX) * m_ScaleX, (cY) * m_ScaleY, (cX) * m_ScaleX, (cY + 1) * m_ScaleY)
CombineRgn MainRGN, MainRGN, FlyRegion, RGN_MAX
DeleteObject FlyRegion 'deleting temporary
End If
bX = -1 'allow to start again, so reseting values
End If
End If
Next cX 'resume horizontally
Next cY 'when finish horizontally, start horizontally again with the next vertical pixel
'finally we craete the last temporary region
FlyRegion = CreateRectRgn(0, 0, srcWidth * m_ScaleX, SrcHeight * m_ScaleY)
'attach it to the main region
CombineRgn MainRGN, FlyRegion, MainRGN, RGN_DIFF
DeleteObject FlyRegion 'free up resources
m_CurrentRGN = MainRGN 'setting our main region
RaiseEvent BitMapLoaded 'raising event
End Function
Public Property Get ScaleX() As Variant
    ScaleX = m_ScaleX
End Property
Public Property Let ScaleX(ByVal New_ScaleX As Variant)
    m_ScaleX = CLng(New_ScaleX)
    PropertyChanged "ScaleX"
End Property
Public Property Get ScaleY() As Variant
    ScaleY = m_ScaleY
End Property
Public Property Let ScaleY(ByVal New_ScaleY As Variant)
    m_ScaleY = CLng(New_ScaleY)
    PropertyChanged "ScaleY"
End Property
Public Property Get MyObject() As Object
Attribute MyObject.VB_MemberFlags = "400"
    Set MyObject = m_MyObject
End Property
Public Property Set MyObject(ByVal New_MyObject As Object)
    If Ambient.UserMode = False Then Err.Raise 383
    Set m_MyObject = New_MyObject
    PropertyChanged "MyObject"
End Property
'the propert bit color is the transpancy color
'means it's the color that we will take from the given
'image when we need to craete region
'we wil make regions contains that color in the image transparent in the window
Public Property Get BitColor() As OLE_COLOR
    BitColor = m_BitColor
End Property
Public Property Let BitColor(ByVal New_BitColor As OLE_COLOR)
    m_BitColor = New_BitColor
    PropertyChanged "BitColor"
End Property
Private Sub UserControl_InitProperties()
    m_CurrentRGN = m_def_CurrentRGN
    m_ScaleX = m_def_ScaleX
    m_ScaleY = m_def_ScaleY
    m_BitColor = m_def_BitColor
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_CurrentRGN = PropBag.ReadProperty("CurrentRGN", m_def_CurrentRGN)
    m_ScaleX = PropBag.ReadProperty("ScaleX", m_def_ScaleX)
    m_ScaleY = PropBag.ReadProperty("ScaleY", m_def_ScaleY)
    Set m_MyObject = PropBag.ReadProperty("MyObject", Nothing)
    m_BitColor = PropBag.ReadProperty("BitColor", m_def_BitColor)
End Sub
Private Sub UserControl_Terminate()
DeleteRGN
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CurrentRGN", m_CurrentRGN, m_def_CurrentRGN)
    Call PropBag.WriteProperty("ScaleX", m_ScaleX, m_def_ScaleX)
    Call PropBag.WriteProperty("ScaleY", m_ScaleY, m_def_ScaleY)
    Call PropBag.WriteProperty("MyObject", m_MyObject, Nothing)
    Call PropBag.WriteProperty("BitColor", m_BitColor, m_def_BitColor)
End Sub
