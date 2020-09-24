VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWordPasswordRecovery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Word Password Recovery Tool"
   ClientHeight    =   2820
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   5550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmWordPasswordRecovery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1988
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dictonary Attack"
      Height          =   375
      Left            =   1988
      TabIndex        =   4
      Top             =   1736
      Width           =   1575
   End
   Begin VB.TextBox FName 
      Height          =   375
      Left            =   304
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   713
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Choose File"
      Height          =   375
      Left            =   188
      TabIndex        =   2
      Top             =   1736
      Width           =   1575
   End
   Begin VB.TextBox Password 
      Height          =   375
      Left            =   1751
      TabIndex        =   1
      Top             =   1193
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Brute Force Attack"
      Height          =   375
      Left            =   3788
      TabIndex        =   0
      Top             =   1736
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog OpenFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "Xls"
      DialogTitle     =   "Choose Excel File to Find Password"
   End
   Begin VB.Label Label1 
      Caption         =   "Password  :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   311
      TabIndex        =   6
      Top             =   1193
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Speed 
      Height          =   495
      Left            =   315
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmWordPasswordRecovery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'PROGRAM    :   Word Password Recovery Tool
'AUTHOR     :   Vikas Madaan
'  __         __        ___      ___
'  \ \       / /       |   \    /   |
'   \ \     / /        | |\ \  / /| |
'    \ \   / /         | | \ \/ / | |
'     \ \_/ /    __    | |  \__/  | |
'      \___/    (__)   |_|        |_|
'
'DATE       :   November 18, 2003.
'
'COMMENTS   :   This is an Word File Password Recovery Tool.
'           It is used to recover password from the Word File.
'           It show the usage of Dictionary Attack &
'           Brute Force Attack upto 2 Character Length
'           But you can increase it to any length.
'           when U modify this code & add New Features
'           then please also send me the copy of that.
'           USE FOR EDUCATIONAL PURPOSES ONLY!!!
'           If you need support or to give suggestions to improve,
'           you can email me at vikasmadaan25@hotmail.com
'           or thru yahoo messenger vikasmadaan25@yahoo.com
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Option Explicit
Dim Char(1 To 62) As String * 1 'Character for Brute Force
Dim tm As Date 'For Total Time
Dim PCount As Long ' To Check Total Password Checked
Dim PLast As Long 'To Check Last Total
Dim Doc As Document  'For Word Document
Dim Pass As String 'Hold the Current Password Applied
Const Title As String = "Word File Password Recovery"
Dim Tmr As Single

'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'This is the main function that checks for the password on
'Word file it Returns True if Password Found.
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Private Function FindPassword(ByVal Pass As String) As Boolean
On Error GoTo NotFound
PCount = PCount + 1
DoEvents
'Open File with the Password
'It return False If Not Open means Password not Valid
Set Doc = Word.Documents.Open(FName.Text, , True, , Pass)
Doc.Close False
FindPassword = True
Exit Function

NotFound:
FindPassword = False
End Function

'Dictionary Btn
Private Sub Command1_Click()
On Error GoTo ErrReadingFile
Label1.Visible = False
Password.Visible = False
'Check For File Selected
If Len(FName.Text) = 0 Then
 MsgBox "No File Selected....." & vbCrLf & "Select The File First.....", vbCritical, Title
 Exit Sub
End If
Dim Find As Boolean
'Check for the file is password protected or not
Find = FindPassword(" ")
If Find Then
 MsgBox "No Password Set For The File" & vbCrLf & "You Can Open File Without Any Password", vbExclamation, Title
 Exit Sub
End If
'Dictionary Attack
'Open Dictionary file to Retrive Words
Open App.Path & "\" & "English.dic" For Input As #1
PCount = PLast = 0
Timer1.Enabled = True
tm = Now
Do Until EOF(1)
 DoEvents
 Line Input #1, Pass
 Find = FindPassword(Pass)
 If Find Then Exit Do
Loop
Timer1_Timer
If Find Then
 MsgBox "Password Found" & vbCrLf & vbCrLf & "Password=""" & Pass & """", , Title
 Password.Text = Pass
 Password.Visible = True
 Label1.Visible = True
Else
 MsgBox "Sorry! Password Not Found", , Title
End If

ErrReadingFile:
Timer1.Enabled = False
Close #1
If Err Then
 MsgBox Err.Description, vbCritical, Title
End If
End Sub

'Choose File Btn
Private Sub Command2_Click()
On Error GoTo Cancel
OpenFile.FileName = ""
OpenFile.Filter = "Word Files (*.Doc)|*.Doc"
OpenFile.Flags = cdlOFNLongNames Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
OpenFile.ShowOpen
FName.Text = OpenFile.FileName
Password.Visible = False
Label1.Visible = False
Exit Sub
Cancel:
End Sub

'Brute Force Btn
Private Sub Command3_Click()
On Error GoTo ErrReadingFile
Label1.Visible = False
Password.Visible = False
'Check for file Selected
If Len(FName.Text) = 0 Then
 MsgBox "No File Selected....." & vbCrLf & "Select The File First.....", vbCritical, Title
 Exit Sub
End If
Dim Find As Boolean, i As Integer, j As Integer
'Check for Password Protected
Find = FindPassword(" ")
If Find Then
 MsgBox "No Password Set For The File" & vbCrLf & "You Can Open File Without Any Password", vbExclamation, Title
 Exit Sub
End If
PCount = PLast = 0
Timer1.Enabled = True
tm = Now
'Brute Force Attack For 2 Charcters
For i = 1 To 62
 For j = 1 To 62
  DoEvents
  Pass = Char(i) & Char(j)
  Find = FindPassword(Pass)
  If Find Then Exit For
 Next
 If Find Then Exit For
Next
Timer1_Timer
If Find Then
 MsgBox "Password Found" & vbCrLf & vbCrLf & "Password=""" & Pass & """", , Title
 Password.Text = Pass
 Password.Visible = True
 Label1.Visible = True
Else
 MsgBox "Sorry! Password Not Found", , Title
End If

ErrReadingFile:
Timer1.Enabled = False
Close #1
If Err Then
 MsgBox Err.Description, vbCritical, Title
End If
End Sub

'Exit Btn
Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
'U can also add any number of Characters
Set Doc = New Document
Dim i As Integer, j As Integer
j = 1
For i = Asc("a") To Asc("z")
 Char(j) = Chr(i)
 j = j + 1
Next i
For i = Asc("A") To Asc("Z")
 Char(j) = Chr(i)
 j = j + 1
Next i
For i = Asc("0") To Asc("9")
 Char(j) = Chr(i)
 j = j + 1
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Doc.Close False
Set Doc = Nothing
End
End Sub

Private Sub Timer1_Timer()
Speed.Caption = "Speed/Sec = " & PCount - PLast & "       Time = " & Format$(Now - tm, "hh:mm:ss") & vbCrLf & "Total = " & PCount & "       Current Password = " & Pass
PLast = PCount
End Sub

