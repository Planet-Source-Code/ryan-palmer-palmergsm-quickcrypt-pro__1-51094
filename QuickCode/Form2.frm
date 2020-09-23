VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim MyString, MyNumber
Open "Testcrypt.txt" For Input As #1 ' Open file for input.
Do While Not EOF(1) ' Loop until end of file.
    Input #1, MyString
    Form2.Text1 = MyString
Loop
Close #1    ' Close file.
End Sub
