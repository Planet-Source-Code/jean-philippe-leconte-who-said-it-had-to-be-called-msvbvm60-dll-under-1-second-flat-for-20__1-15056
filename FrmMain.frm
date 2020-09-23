VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Patch EXE Runtime dll 3 - Modified by Jean-Philippe Leconte"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.TextBox OldDll 
         Height          =   285
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   1
         Text            =   "MSVBVM60.DLL"
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox DLLName 
         Height          =   285
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox FileToPatch 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Patch"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This version enables you to patch the EXE in a matter of seconds!!!"
         Height          =   390
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   3375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old DLL Required:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DLL to use:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What would you like the EXE Runtime dll to be called?"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXE to patch:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   990
      End
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code by Mark Withers
' Modified by Jean-Philippe Leconte
' In no way I wanted to steal this code
' I just wanted to show how to optimize speed in VB

' Hope it helps you Mark

Private Sub Command3_Click()
    Dim sBuffer As String * 1024 ' Buffer
    Dim BufferLength As Long ' Buffer length
    Dim FileNumber As Integer ' File number
    Dim FileLength As Long ' File length
    Dim FilenameLength As Long ' filename length
    Dim sNewDLL As String ' NEW DLL Name with null char terminating
    Dim Pos As Long ' Current pos in file
    Dim PosFound As Long ' Pos found the dll name
    
    ' Validation
    If DLLName.Text = "" Then MsgBox "Please enter a new DLL name for your program.", vbOKOnly, "Patch": Exit Sub
    If OldDll.Text = "" Then MsgBox "Please enter the runtime DLL required for you program.", vbOKOnly, "Patch": Exit Sub
    If FileToPatch.Text = "" Then MsgBox "Please enter a EXE file to patch in the EXE input box.", vbOKOnly, "Patch": Exit Sub
    If FileLen(FileToPatch.Text) <= 0 Then MsgBox "Please enter a valid EXE file to patch into the EXE input box.", vbOKOnly, "Patch": Exit Sub
    If Len(OldDll.Text) < Len(DLLName.Text) Then MsgBox "Please enter a new DLL name shorter or equal to the old one.", vbOKOnly, "Patch": Exit Sub
    
    ' Calculate length before looping, reduce calculation time A LOT!
    BufferLength = Len(sBuffer)
    sNewDLL = DLLName.Text + String(Len(OldDll.Text) - Len(DLLName.Text), vbNullChar)
    FilenameLength = Len(OldDll.Text)
    FileLength = FileLen(FileToPatch.Text)
    Pos = 1
    
    ' Open file
    FileNumber = FreeFile
    Open FileToPatch.Text For Binary As FileNumber
        While Not EOF(FileNumber)
            ' Get X chars from file (in this case 1024, but you could change de 1024 to anything else...
            ' the greater the number, the more memory, but the faster is runs)
            Get FileNumber, Pos, sBuffer
            PosFound = InStr(sBuffer, OldDll.Text) ' Search DLL name
            ' If found, write new DLL name
            If PosFound > 0 Then
                Put FileNumber, Pos + PosFound - 1, sNewDLL
            End If
            ' Check next set of X chars, but remove the filename length in case
            ' it was truncated, (in case we have only the first 4 chars at the end of the buffer)
            Pos = Pos + BufferLength - FilenameLength
        Wend
    Close FileNumber
    
    MsgBox "The new Dll for Runtime has been written into the EXE.", vbOKOnly, "Patch"
End Sub
