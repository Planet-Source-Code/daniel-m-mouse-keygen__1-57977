VERSION 5.00
Begin VB.Form frmGenerate 
   Caption         =   "Mouse Random Key Generator"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAllow 
      Caption         =   "Allow Numeric Characters"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CheckBox chkAllow 
      Caption         =   "Allow Lower-case Alpha Characters"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox chkAllow 
      Caption         =   "Allow Upper-case Alpha Characters"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.Timer tmrGData 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2640
      Top             =   1680
   End
   Begin VB.TextBox txtRKey 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   3960
      TabIndex        =   3
      Text            =   "000-000-000-000-000"
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Start Generator"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblNotice 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmGenerate.frx":0000
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3360
      TabIndex        =   8
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label lblGTime 
      Caption         =   "Generation Time: 0 seconds"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblRKey 
      Caption         =   "Random Key:"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblDesc 
      Caption         =   $"frmGenerate.frx":00D7
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmGenerate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
'Author: Daniel M.
'Comments: Somebody gave me a comment about creating a key generator based on
'mouse movement and I had already heard about this previously and decided to
'make my own version of it. Well here you have it. You can modify this key-gen
'to be more random if you'd like, it is just a simple example. I'm too lazy to
'make more settings modifiable such as interval for recording, choosing save
'location, whatever else.. But anyhow, please VOTE FOR ME!
'
'Extra Notes: AT PSC search for "DNS Browser" for a 100+ source code web browser
'I spent over 2 1/2 months working on it so go check it out and give me feedback
'===============================================================================

Option Explicit
Private Declare Function GetCursorPos Lib "user32" (lppoint As POINTAPI) As Long 'mouse declaration for retrieving info
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Type POINTAPI 'set type
    X As Long
    Y As Long
End Type

'Declarations
Dim gTime As Long, gCount As Long, gKey As String
Dim Cursor As POINTAPI
Private Sub cmdGenerate_Click()
Dim fso 'File System Object for file checking
Set fso = CreateObject("Scripting.FileSystemObject")

gCount = 0
gTime = 0
gKey = vbNullString

If cmdGenerate.Caption = "Start Generator" Then
    If fso.FileExists(App.Path & "\rnd_data.dat") Then 'if exists, we must create new data
        If MsgBox("Temporary data file already exists, delete it?", vbYesNo + vbCritical, _
        "Temporary File Exists") = vbYes Then
            Kill App.Path & "\rnd_data.dat" 'if exists and ok'd, delete old file
        Else
            startGeneration 'start function
            Exit Sub
        End If
    End If
    
    tmrGData.Enabled = True 'start generating random data
    cmdGenerate.Caption = "Stop Generator"
Else
    tmrGData.Enabled = False 'stop generating random data
    cmdGenerate.Caption = "Start Generator"
    startGeneration
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then Call cmdGenerate_Click
End Sub

Private Sub tmrGData_Timer()
If GetAsyncKeyState(vbKeyEnd) Then Call cmdGenerate_Click
Dim rndPlacement As Integer 'make var

Randomize 'create randomization
rndPlacement = Int(Rnd * 2) + 1

GetCursorPos Cursor

Open App.Path & "\rnd_data.dat" For Append As #1
    If rndPlacement = 1 Then 'we use this to make more randomness
        Print #1, Cursor.X & Cursor.Y
    Else
        Print #1, Cursor.Y & Cursor.X
    End If
Close #1

gCount = gCount + 1 'count time

If gCount = 100 Then 'if 100 revolutions then 1 second has passed
    gTime = gTime + 1
    gCount = 0
    lblGTime.Caption = "Generation Time: " & gTime & " seconds"
End If


End Sub
Private Function startGeneration()
Dim rndData As String, tempStr As String, lngRndNum As Long, lngNum As Long, chrNum As String, lngReverse As Long

Open App.Path & "\rnd_data.dat" For Input As #1 'open file to gather generated data
    Do While Not EOF(1)
        Input #1, tempStr$
        rndData$ = rndData$ & tempStr$
    DoEvents
    Loop
Close #1

Do Until Len(gKey$) >= 15 'do until the key length is 15, you can set this to whatever you like
    Randomize 'generate a random
    lngRndNum = Int(Rnd * Len(rndData$)) + 1 'used to get location of random number
    
    Randomize
    lngReverse = Int(Rnd * 2) + 1
    
    If lngReverse = 1 Then 'Produces even more randomness by possibly randomly reversing numbers
        lngNum = StrReverse(Mid$(rndData$, lngRndNum, 2))
    Else
        lngNum = Mid$(rndData$, lngRndNum, 2)
    End If

    'checks for restrictions
    If chkAllow(2).Value = 1 And lngNum >= 48 And lngNum <= 57 Or chkAllow(0).Value = 1 And lngNum >= 65 And _
    lngNum <= 90 Or chkAllow(1).Value = 1 And lngNum >= 10 And lngNum <= 35 Then 'we use 10-35 instead of 97-122
        If lngNum <= 35 Then 'since we only use 2 digit numbers, we must define lowercase chars as different numbers
            chrNum$ = Chr(lngNum + 87)
            gKey = gKey$ & chrNum$
        Else
            chrNum$ = Chr(lngNum) 'if the number is an ascii value then convert to the chr if result is alphanumeric
            gKey$ = gKey$ & chrNum$
        End If
    Else
        'do nothing, keep looping until key is generated
    End If
    
    'uncomment the line below to see progress in debugger
    'Debug.Print gKey$
DoEvents
Loop

Dim i As Long, fKey As String 'fKey is the final key with "-"

For i = 1 To Len(gKey$)
    fKey$ = fKey$ & Mid(gKey$, i, 1)
    If i = Len(gKey$) Then Exit For
    If i Mod 3 = 0 Then fKey = fKey$ & "-"
DoEvents
Next i

txtRKey.Text = fKey$ 'show generated key
txtRKey.SetFocus
txtRKey.SelLength = Len(txtRKey)
End Function
