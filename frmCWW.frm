VERSION 5.00
Begin VB.Form frmCWW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmCWW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.OptionButton optCWW 
      Caption         =   "8 Hour Days - Standard"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton optCWW 
      Caption         =   "9 Hour Days - Compressed"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Compressed Work Week Option:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmCWW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    ' Saving the Week Type settings for calculation purposes
    WriteINI "CWW", "CWW", strOptCWW, App.Path & "\Logtime.ini"
'    Unload Me
    Call Form_Load
End Sub

Private Sub Form_Load()
On Error Resume Next
    'Centering the form on screen
    'CenterMe Me
    Me.Top = dblMainTop + 3105
    Me.Left = dblMainLeft + 5010
    DockingStart Me
    ' This reads the current setting from the ini file and displays
    ' it to the user
    x = ReadINI("CWW", "CWW", App.Path & "\Logtime.ini")
    If x = "1" Then
        optCWW(1).Value = True
    Else
        optCWW(0).Value = True
    End If
    Me.Caption = "Set To:" & optCWW(x).Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'DockingTerminate Me
End Sub

Private Sub optCWW_Click(Index As Integer)
    ' The selection here stores the week type settings is a variable for
    ' later use
    Select Case Index
        Case 0
            strOptCWW = "0"
        Case 1
            strOptCWW = "1"
    End Select
End Sub
