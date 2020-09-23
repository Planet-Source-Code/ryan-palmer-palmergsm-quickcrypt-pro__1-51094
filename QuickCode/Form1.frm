VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "QuickCrypt Lite : Version 0.5B"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8040
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "EasyDecrypt"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   7575
      Begin VB.TextBox txtuntimes 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   18
         Text            =   "1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Decryption Level:"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox decryptionkey 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   1320
         TabIndex        =   14
         Text            =   "Encyc$"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Text            =   "Encryption Key:"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Text to be decrypted"
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label clrtext2 
         BackColor       =   &H80000007&
         Caption         =   "Clear"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   6840
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.Label command2 
         BackColor       =   &H80000007&
         Caption         =   "Decrypt"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   6840
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Encryption Level:"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txttimes 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   4800
      MaxLength       =   2
      TabIndex        =   15
      Text            =   "1"
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "EasyResult"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   7575
      Begin VB.TextBox Text3 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF8080&
         Height          =   405
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "Form1.frx":0442
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Caption         =   "Clear"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   6840
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "EasyCrypt"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   7575
      Begin VB.TextBox encryptionkey 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   1320
         TabIndex        =   13
         Text            =   "Encyc$"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox lbltext 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Encryption Key:"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "Text to be encrypted"
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label clrtext1 
         BackColor       =   &H80000007&
         Caption         =   "Clear"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   6840
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Command1 
         BackColor       =   &H80000007&
         Caption         =   "Encrypt"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   5160
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004000&
      X1              =   240
      X2              =   7800
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "QuickCrypt - Instant Text Encryption. File and Document version coming soon"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PalmerGSM 2003-2004 - Designed and Made by Ryan Palmer. Version 0.5 ; Beta release testing stage"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4800
      Width           =   7575
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C00000&
      X1              =   7920
      X2              =   7920
      Y1              =   4920
      Y2              =   1320
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C00000&
      X1              =   120
      X2              =   7920
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF0000&
      X1              =   120
      X2              =   120
      Y1              =   5040
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   120
      X2              =   7920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Image Image1 
      Height          =   1605
      Left            =   1080
      Picture         =   "Form1.frx":0453
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clrtext1_Click()
Form1.Text1.Text = ""
End Sub

Private Sub clrtext2_Click()
Form1.Text2.Text = ""
End Sub

'Beta code, not subbed very well. By Ryan Palmer
Private Sub Command1_Click()

Form1.Text3.Text = ""
Dim snozza As Integer
Dim scuzz
If txttimes = 0 Then Exit Sub

scuzz = Text1.Text
Dim text3save1
Do
If Not snozza = 0 Then text3save1 = Text3.Text
If Not snozza = 0 Then Text1.Text = Text3.Text
If Not snozza = 0 Then Text3.Text = ""

Form1.ProgressBar1.Value = ((100 / txttimes) * snozza)
snozza = snozza + 1
On Error GoTo ErrorHandler
Dim n
Dim p
Dim bobba
Dim shrub As Integer
shrub = "2"
n = 1
Do

DoEvents
p = Mid$(Form1.Text1.Text, n, 1) 'Get the character to work on
If Not leavealone = "1" Then p = Asc(p) 'Convert to ASCII code
Dim c As Integer
c = p
c = c + shrub '''''''''''''''''''''''''''''''''''
shrub = shrub + 2
If shrub = "20" Then shrub = "2"

c = c
p = c
leavealone = "0"
'MsgBox p 'Debug Purposes
If Len(p) = 3 Then
Dim MyValue
MyValue = Int((25 * Rnd) + 1)    ' Generate random value between 1 and 6.
bobba = Chr(MyValue + 65)
End If
If Len(p) = 2 Then
Dim MyValue2
MyValue2 = Int((20 * Rnd) + 1)    ' Generate random value between 1 and 6.
bobba = Chr(MyValue2 + 65) + Chr(MyValue2 + 70)
End If
Dim MyValue3
MyValue3 = Int((9 * Rnd) + 1)    ' Generate random value between 1 and 6.
Form1.Text3.Text = Form1.Text3.Text & p & bobba & MyValue3 'Say that whatever needs doing has been done
n = n + 1 'Next please
Loop Until n = Len(Form1.Text1.Text) + 1 'Until the length has finished
yollo = Asc(Form1.encryptionkey)
yollo = yollo + 2
Form1.Text3.Text = Form1.Text3.Text & yollo
'MsgBox "Once round"
  Loop Until snozza = txttimes
                Text1.Text = scuzz
                Form1.ProgressBar1.Value = 100
ErrorHandler:   ' Error-handling routine.
    Select Case Err.Number  ' Evaluate error number.
        Case 16 ' Type mismatch
            MsgBox "Oh errrm. some sort of error"   ' Close open file.
            Case 13 'Oh no
            MsgBox "Some sort of validation error occured"
            Case 5
            MsgBox "A massive runtime error. Possible bad character code. Try removing strange characters"
                    Case Else
            If Not Err.Number = 20 Then If Not Err.Number = 0 Then MsgBox "Holy Crap - You got error code : " & Err.Number

End Select
    Exit Sub  ' Resume execution at next line
                ' that caused the error.
              
End Sub


Private Sub Command2_Click()
Form1.Text3.Text = ""

If txtuntimes = 0 Then Exit Sub
Dim snubbo
snubbo = Text2.Text
Dim snubby As Integer

Do
If Not snubby = 0 Then Text2.Text = Text3.Text
If Not snubby = 0 Then Text3.Text = ""
Form1.ProgressBar1.Value = ((100 / txtuntimes) * snubby)
snubby = snubby + 1
Dim spx As Integer
spx = "2"
Dim n
Dim p
Dim lenp
Dim lenpx
Dim Doomed
On Error GoTo ErrorHandler
n = 1
p = "dx"
Dim abba
Dim baab
abba = Asc(Form1.decryptionkey)
Dim lendy
lendy = Len(Form1.Text2.Text) - 1
baab = Right$(Form1.Text2.Text, 2)
abba = Chr(abba)
baab = Chr(baab - 2)

If Not abba = baab Then
MsgBox "Wrong Decryption Key, Decryption Aborted"
Doomed = "1"
Else
Form1.Text2.Text = Mid$(Form1.Text2.Text, 1, lendy)
End If
If Doomed = "1" Then Form1.Text3.Text = "<<<ABORTED DECRYPTION>>>"
Do
If p = "" Then Exit Sub
p = Mid$(Form1.Text2.Text, n, 5)
Dim wuzza
Dim done
lenp = Len(p) - 1
p = Left$(p, lenp)
Do
DoEvents
'Purposely takes time so the user aint disappointed
wuzza = wuzza + 1
If Right$(p, 1) = "0" Then done = "1"
If Right$(p, 1) = "1" Then done = "1"
If Right$(p, 1) = "2" Then done = "1"
If Right$(p, 1) = "3" Then done = "1"
If Right$(p, 1) = "4" Then done = "1"
If Right$(p, 1) = "5" Then done = "1"
If Right$(p, 1) = "6" Then done = "1"
If Right$(p, 1) = "7" Then done = "1"
If Right$(p, 1) = "8" Then done = "1"
If Right$(p, 1) = "9" Then done = "1"
If Right$(p, 1) = Null Then done = "1"
If Right$(p, 1) = "" Then done = "1"
If Not done = "1" Then lenp = Len(p) - 1
If Not done = "1" Then p = Left$(p, lenp)
If wuzza = "5" Then done = "1"
Loop Until done = "1"
done = "0"
'MsgBox p <-- Debug Purposes
If p = "" Then fuckoff = "1"
Dim g
If Doomed = "1" Then Form1.Text3.Text = "<<<ABORTED DECRYPTION>>>"
If Not fuckoff = "1" Then g = p

If spx = Null Then spx = "2"
If spx = "20" Then spx = "2"
If Not fuckoff = "1" Then g = g - spx
On Error GoTo ErrorHandler
If Not fuckoff = "1" Then p = Chr(g)



Form1.Text3.Text = Form1.Text3.Text + p
spx = spx + 2
n = n + 5
If Doomed = "1" Then Form1.Text3.Text = "<<<ABORTED DECRYPTION>>>"
If Doomed = "1" Then Form1.Text3.Text = "<<<ABORTED DECRYPTION>>>"
Doomed = "0"
Loop Until n = Len(Form1.Text2.Text)
Loop Until snubby = txtuntimes
Form1.ProgressBar1.Value = 100
ErrorHandler:   ' Error-handling routine.
    Select Case Err.Number  ' Evaluate error number.
        Case 16 ' Type mismatch
            MsgBox "Decryption error, attempted code is not compatible with ASCII"   ' Close open file.
        Case Else
            If Not Err.Number = 20 Then If Not Err.Number = 0 Then If Not Err.Number = 5 Then MsgBox "Holy Crap - You got error code : " & Err.Number
If Not Err.Number = 20 Then If Not Err.Number = 0 Then Exit Sub
End Select
    
End Sub

Private Sub Command3_Click()
Form2.Show
End Sub

Private Sub Form_Load()
MsgBox "QuickCrypt Lite - NOTICE FOR USERS: - Please note that the encryption level is on a quadratic curve, and any value above 4 is likely to take some time, and makes the key a lot longer.)"
Form1.ProgressBar1.Value = "100"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_Click()
Text3.Text = ""
'Cleared - wow!
End Sub


Private Sub Label4_Click()

End Sub
