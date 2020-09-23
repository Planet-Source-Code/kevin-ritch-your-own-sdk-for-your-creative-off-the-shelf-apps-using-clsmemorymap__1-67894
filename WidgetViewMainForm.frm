VERSION 5.00
Begin VB.Form WidgetViewMainForm 
   Caption         =   "WidgetView - A useless thingy to demonstrate SDK Development"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   Icon            =   "WidgetViewMainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton UselessButton 
      Caption         =   "Click Me"
      Height          =   615
      Left            =   4800
      TabIndex        =   5
      Top             =   720
      Width           =   2295
   End
   Begin VB.Timer MySDK 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox PeopleList 
      Height          =   1035
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label RequestStr 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6435
      TabIndex        =   3
      Top             =   285
      Width           =   615
   End
   Begin VB.Label ResultLabel 
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   420
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   2025
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User selected:"
      Height          =   1095
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   2280
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add On Requested"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   240
      Width           =   2265
   End
End
Attribute VB_Name = "WidgetViewMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TheRequest As New clsMemoryMap
Dim TheResult As New clsMemoryMap
Private Sub Form_Load()
'====================
'Open the Memory Maps
'====================
 TheRequest.OpenMemory ("VB6Request")
 TheResult.OpenMemory ("VB6Result")
'=======================
'Set the List with stuff
'=======================
 PeopleList.AddItem "Bob"
 PeopleList.AddItem "Carol"
 PeopleList.AddItem "Ted"
 PeopleList.AddItem "Alice"
 PeopleList.AddItem "Kevin"
 PeopleList.ListIndex = PeopleList.ListCount - 1 ' (Last one)
 DoEvents
'================
'Flush the values
'================
 TheRequest.Poke ""
 TheResult.Poke ""
 MySDK.Enabled = True
End Sub
Private Sub PeopleList_Click()
 MySDK.Enabled = False
 ResultLabel = PeopleList
 RequestStr = Trim$(PeopleList.ListIndex + 1)
 TheRequest.Poke RequestStr
 MySDK.Enabled = True
End Sub
Private Sub RequestStr_Change()
 MySDK.Enabled = False
 Kommand$ = RequestStr
 Select Case UCase$(Trim$(Kommand$))
  Case "1", "2", "3", "4", "5"
  '=================================================================
  'THE ADD-ON APPLICATION APPARENTLY WANTS TO SEE A NAME IN THE LIST
  '=================================================================
   ThePerson = Val(Kommand$)
   PeopleList.ListIndex = ThePerson - 1
   PeopleList.Refresh
  Case Else
  'Whatever
 End Select
 MySDK.Enabled = True
End Sub
Private Sub MySDK_Timer()
 RequestStr = TheRequest.Peek
End Sub
Private Sub ResultLabel_Change()
 TheResult.Poke ResultLabel
End Sub

Private Sub UselessButton_Click()
 MsgBox "Hello World Wide Web!" & String$(2, 10) & "OK, so this button is here just to remind you all about voting." & String$(2, 10) & "Thanks!", vbApplicationModal + vbInformation, "Thanks for pressing this button"
End Sub
