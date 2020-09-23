VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5505
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me Please"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Trans1.ScaleX = 2.5 'Set The Region to be large
Form1.Trans1.ScaleY = 1.5 'Set the region to be large vertically
Form1.Trans1.LoadBitmap App.Path & "\Trans2.bmp" ' the shape to apply to text1 of form2 ... see how
Set Form1.Trans1.MyObject = Form2.Text1
Form1.Trans1.DoObject 'woow the text box is  ...!!! wonderful
Form1.Trans1.DeleteRGN 'free up resources
'Note , You must set the scaleX , and ScaleY before loading the bitmap
End Sub
