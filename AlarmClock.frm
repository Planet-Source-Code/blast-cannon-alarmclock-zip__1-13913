VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alarm Clock"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2175
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AlarmClock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Set Alarm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1470
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   -15
      Top             =   15
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8888"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   -15
      TabIndex        =   1
      Top             =   45
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AlarmTime As Variant    'The AlarmTime can be in 12- or 24- hour form, but 24- hour is recommended because it is easier to work with
Dim Message As String       'The message you want to remember stored as text
Private Sub Command1_Click()    'Clicking "Set Alarm"
    AlarmTime = InputBox("Enter alarm time:", "Alarm Settings", AlarmTime)  'Displays the Input box with the AlarmTime box
    Message = InputBox("Note to remember:", "Note", Message)    'Displays the Note To Remeber box
    If AlarmTime = "" Then Exit Sub     'If the user did not enter a time
    If Not IsDate(AlarmTime) Then   'If the user did not enter a valid time (such as they typed words or letters)
        MsgBox "Not a valid time."  'a Message Box is displayed
    Else
        AlarmTime = CVDate(AlarmTime)   'Converts the entered time into system time (day/month/year or its equivalent)
    End If
End Sub

Private Sub Timer1_Timer()
Static AlarmSounded As Integer  'Allows for the application to continue running and beep even if it doesn't have the focus
    If Label1.Caption <> CStr(Time) Then    'If the time display is not the same as the System Clock...
        Label1.Caption = Time   'then it sets the display equal to the System Time
        If Time >= AlarmTime And Not AlarmSounded Then    'If the System Time is past the Alarm Time then...
            Beep    'The computer beeps and...
            MsgBox Message  'the Note To Remeber is displayed
            AlarmSounded = True 'Sets the AlarmSounded proplerty to True so the next statement does not interfere with normal operation
        ElseIf Time < AlarmTime Then    'If the Alarm Time has not yet been reached...
            AlarmSounded = False    'Then the AlarmSounded property is False because it has not been reached yet
    End If
End If
End Sub


