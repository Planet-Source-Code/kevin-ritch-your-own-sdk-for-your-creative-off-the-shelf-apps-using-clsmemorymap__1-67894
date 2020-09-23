VERSION 5.00
Begin VB.Form MyAddOnMainForm 
   Caption         =   "MyAddOn for WidgetView"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   Icon            =   "MyAddOnMainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton PersonNumberWanted 
      Caption         =   "Ask for #5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GET LATEST DATA ONLY"
      Height          =   615
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
   Begin VB.OptionButton PersonNumberWanted 
      Caption         =   "Ask for #4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton PersonNumberWanted 
      Caption         =   "Ask for #3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.OptionButton PersonNumberWanted 
      Caption         =   "Ask for #2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.OptionButton PersonNumberWanted 
      Caption         =   "Ask for #1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton RequestDataButton 
      Caption         =   "REQUEST SELECTED DATA"
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Timer MySDK 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   1680
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
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   2025
   End
End
Attribute VB_Name = "MyAddOnMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RequestStr As String
Dim TheRequest As New clsMemoryMap
Dim TheResult As New clsMemoryMap

Private Sub Command1_Click()
 a$ = TheResult.Peek
 s = InStr(a$, Chr$(0))
 a$ = IIf(s > 0, Left$(a$, s - 1), a$)
 ResultLabel = a$
End Sub

Private Sub Form_Load()
'====================
'Open the Memory Maps
'====================
 TheRequest.OpenMemory ("VB6Request")
 TheResult.OpenMemory ("VB6Result")
End Sub
Private Sub MySDK_Timer()
 a$ = TheResult.Peek
 s = InStr(a$, Chr$(0))
 a$ = IIf(s > 0, Left$(a$, s - 1), a$)
 ResultLabel = a$
End Sub

Private Sub RequestDataButton_Click()
 ResultLabel = ""
 For i = 0 To 4
  If PersonNumberWanted(i) Then
   RequestStr = Trim$(i + 1)
   Exit For
  End If
 Next i
 RequestStr = RequestStr
 TheRequest.Poke RequestStr
 MySDK.Enabled = True
End Sub

Private Sub ResultLabel_Change()
 MySDK.Enabled = False
End Sub
