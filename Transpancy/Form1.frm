VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.Trans Trans1 
      Left            =   1800
      Top             =   600
      _ExtentX        =   2143
      _ExtentY        =   1931
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Project Show How To Use The Transparent Object Contirol in your programs,
'it 's 100% Free , Dosen't Display any Message. so Powerful and Easy To Use
'(c) Marco Samy Nasif, 2002 .. Freeware. you can use,publish it as if you want
'
'This Control Can Trasparent Two Thing (Forms and Objects)
'The Way is to Fix a region by Loading a bitmap, after fixing the BitColor (The Color that will be transparent)
'in seconds the control will extract the Region
'apply this region Using "DoToObject" To a fixed Object
'after the region is not needed Call "DeleteRGN" to free up the memory
Private Sub Form_Load()
Trans1.BitColor = vbWhite 'The White is the color of the transpancy
Set Trans1.MyObject = Me  'The Object to apply Tanspancy using DoObject is This Form
Trans1.LoadBitmap App.Path & "\Trans1.bmp" ' Load the Tansparent shape from the bitmap
'now the shap is kept inside the control
'you can apply it to many forms using DoToObject ( "The Form Name")
Trans1.DoObject 'Now The Form1 (Me) is Tansparent
Form2.Show , Me 'Showing from with WOW
Trans1.DeleteRGN 'free up resources
End Sub
