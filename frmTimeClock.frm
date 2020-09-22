VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTimeClock 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punch Clock"
   ClientHeight    =   4110
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4920
   Icon            =   "frmTimeClock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4920
   Begin MSComctlLib.ImageList ilMenu 
      Left            =   1920
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeClock.frx":0442
            Key             =   "save"
            Object.Tag             =   "save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeClock.frx":0554
            Key             =   "open"
            Object.Tag             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeClock.frx":0666
            Key             =   "date"
            Object.Tag             =   "date"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeClock.frx":0800
            Key             =   "clock"
            Object.Tag             =   "clock"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ilMAIN 
      Left            =   1920
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeClock.frx":099A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeClock.frx":0AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeClock.frx":0BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeClock.frx":1310
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeClock.frx":14AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeClock.frx":15BC
            Key             =   "copy"
            Object.Tag             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeClock.frx":16CE
            Key             =   "paste"
            Object.Tag             =   "paste"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMAIN 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ilMAIN"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Load Date"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Insert Current Time"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Work Week Schedule"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbMAIN 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   3735
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1032
            MinWidth        =   353
            Text            =   "Time:"
            TextSave        =   "Time:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   "10:54 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   926
            MinWidth        =   441
            Text            =   "Date:"
            TextSave        =   "Date:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   "8/23/02"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkLunch 
      Caption         =   "No Lunch"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   18
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Time"
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
      Index           =   1
      Left            =   2760
      TabIndex        =   15
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Time Area"
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
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hours OT:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hours Worked:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Out:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lunch End:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lunch Start:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time In:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "&Data"
      Begin VB.Menu mnuCut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuBreak4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "&Insert Today"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompressed 
         Caption         =   "&Week Type"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTip 
         Caption         =   "&Tip of the Day"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuBreak3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPopRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuPopExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTimeClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------
' Author      David Blaise
' Created     08/13/2002
' Purpose     Stores (specialized) hours worked and lunch hours for time
'             tracking.
' Notes       Incomplete since more validation is needed on data entry.
' Input       Hours starting and hours ending for day and lunch times.
' Output      Hours worked and OT hours worked.
' Revisions   Version 1.0.0: Use of timeclock form, calendar control and
'                            week type.  INI value storage available.
'             Version 1.1.0: Added About, Splash, and Tips along with
'                            minute validation.
'             Version 1.1.1: Added menu bitmaps (looking to imporove)
'             Version 1.2.0: Fixed inability to modify old records,
'                            form binding, cut-copy-paste, and systray icon
'             Version 1.2.1: Added save catch, single day calendar select
'                            for updating old records never started, and
'                            removal of taskbar selection for calendar form
'             Version 1.2.2: Added format to Hours Worked and OT
'             Version 1.3.0: Added Lunch / No Lunch Option & removed mnuicons
'             Version 1.3.1: Fixed Error on open with no INI file
'-----------------------------------------------------------------------------------------

Public Function CalculateLunch(strIn As String, strOut As String) As Date
    Dim dtIn As Date
    Dim dtOut As Date
    
    ' The code may be overdoing the variables here but it is easier for me
    ' to read this way.  Basically the code is formatting each date being
    ' passed to the function into strings
    strIn = Format(strIn, "hh:mm AM/PM")
    strOut = Format(strOut, "hh:mm AM/PM")
    
    ' Now the code convert the strings into dates
    dtIn = CDate(strIn)
    dtOut = CDate(strOut)
    
    ' Here the codes is checking the End of lunch time vs the Start
    ' of lunch time to make sure the user has not entered an end time
    ' greater than the start
    If (dtOut - dtIn) < 0 Then
        MsgBox "The end of your lunch time can not be before the time you left for lunch.", vbOKOnly + vbInformation, "Check Lunch Times"
    Else
    ' This is where the function passes the Lunch time spent back to the
    ' calculation
        CalculateLunch = CDate(Hour(dtOut - dtIn) + (Minute(dtOut - dtIn)) / 60)
    End If
End Function

Public Function CalculateDay(strIn As String, strOut As String) As Date
    ' Ditto to the this function since it is basically the same. It only
    ' varies in the fields it is performing the calculation on
    Dim dtIn As Date
    Dim dtOut As Date
    
    strIn = Format(strIn, "hh:mm AM/PM")
    strOut = Format(strOut, "hh:mm AM/PM")
    
    dtIn = CDate(strIn)
    dtOut = CDate(strOut)
    
    If (dtOut - dtIn) < 0 Then
        MsgBox "The end of your day can not be before the time you left for the day.", vbOKOnly + vbInformation, "Check Day Times"
    Else
        CalculateDay = CDate(Hour(dtOut - dtIn) + (Minute(dtOut - dtIn)) / 60)
    End If
End Function

Private Sub chkLunch_Click()
    If chkLunch.Value = 1 Then
        txtTime(1).Enabled = False
        txtTime(2).Enabled = False
        txtTime(1).Text = Format("12 pm", "HH:MM AM/PM")
        txtTime(2).Text = Format("12 pm", "HH:MM AM/PM")
    Else
        txtTime(1).Enabled = True
        txtTime(2).Enabled = True
        txtTime(1).Text = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(1).Caption, App.Path & "\Logtime.ini")
        txtTime(2).Text = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(2).Caption, App.Path & "\Logtime.ini")
    End If
End Sub

Private Sub cmdCalc_Click()
    Dim x As Integer
    
    ' The code is now validating that each field has a value.  It does not have
    ' data type validation check yet, but this was a start
    For x = 0 To 3
        If txtTime(x).Text = "" Then
            MsgBox "Please enter a " & Label1(x).Caption, vbOKOnly + vbCritical, "No " & Label1(x).Caption & " Entered"
            txtTime(x).SetFocus
            Exit Sub
        End If
    Next x
    
    ' Code is performing the lunch time calculation
    txtTime(4).Text = Format(CalculateDay(txtTime(0), txtTime(3)) - CalculateLunch(txtTime(1), txtTime(2)), "##0.0#")
    ' Code retrieves the Week Type variable and uses it for calculations
    ' It also is verifying that if the day is a Friday under the Compressed
    ' week type, to use a different number.  More functionality will be added
    ' later so the user can select the hours for each day if needed.
    If strOptCWW = "1" Then
        If Format(Now, "dddd") = "Friday" Then
            txtTime(5).Text = Format(CDbl(txtTime(4).Text) - 8, "##0.0#")
        Else
            txtTime(5).Text = Format(CDbl(txtTime(4).Text) - 9, "##0.0#")
        End If
    Else
        txtTime(5).Text = Format(CDbl(txtTime(4).Text) - 8, "##0.0#")
    End If
End Sub

Private Sub cmdReset_Click()
    dtSelected = Now
    ' Resets the data in the text fields
    For Index = 0 To 5
        txtTime(Index).Text = ""
    Next Index
    chkLunch.Value = 0
    Me.Caption = "Punch Clock " & Format(dtSelected, "mm/dd/yyyy")
    txtTime(0).SetFocus
End Sub

Private Sub Form_Load()
    'dtSelected = Now
    ' Here the form loads the variable for the Week Type
    Me.Caption = "Punch Clock " & Format(Now, "mm/dd/yyyy")
    Me.Top = Left(CenterMe(Me), 4)
    Me.Left = Right(CenterMe(Me), 4)
    dblMainTop = Me.Top
    dblMainLeft = Me.Left
    DockingStart Me
    'Week Type variable
    strOptCWW = ReadINI("CWW", "CWW", App.Path & "\Logtime.ini")
    frmCalendar.Show
    frmCWW.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Sys As Long
    Sys = x / Screen.TwipsPerPixelX

    Select Case Sys
        Case WM_LBUTTONDOWN:
            Me.PopupMenu mnuPopUp
        Case WM_RBUTTONDOWN:
            Me.PopupMenu mnuPopUp
'        Case WM_LBUTTONUP:
'            Me.PopupMenu mnuPopUp
'        Case WM_RBUTTONUP:
'            Me.PopupMenu mnuPopUp
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim frm As Form
    
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next frm
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then
        Me.Hide
        frmCalendar.Hide
        frmCWW.Hide
            With nid
                .cbSize = Len(nid)
                .hwnd = Me.hwnd
                .uId = vbNull
                .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
                .uCallBackMessage = WM_MOUSEMOVE
                .hIcon = Me.Icon
                .szTip = Me.Caption & vbNullChar
            End With
        Shell_NotifyIcon NIM_ADD, nid
    Else
        Shell_NotifyIcon NIM_DELETE, nid
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nid
    End
    'DockingTerminate Me
End Sub

Private Sub mnuPopExit_Click()
    Unload Me
End Sub

Private Sub mnuPopRestore_Click()
    WindowState = vbNormal
    Me.Show
    frmCalendar.Show
    frmCWW.Show
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuCompressed_Click()
    frmCWW.Show
End Sub

Private Sub mnuCopy_Click()
    ' copy
    If txtTime(intTextFieldLoc).SelText = "" Then
        Exit Sub
    Else
        Clipboard.Clear
        Clipboard.SetText txtTime(intTextFieldLoc).SelText
    End If
End Sub

Private Sub mnuCut_Click()
    ' cut
    If txtTime(intTextFieldLoc).SelText = "" Then
        Exit Sub
    Else
        Clipboard.Clear
        Clipboard.SetText txtTime(intTextFieldLoc).SelText
        txtTime(intTextFieldLoc).SelText = ""
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuInsert_Click()
    ' This is for setting the current time to text fields.  The code skips
    ' the Hours fields since those are for calculation purposes only
    If intTextFieldLoc > 3 Then
        Exit Sub
    Else
        txtTime(intTextFieldLoc).Text = Format(Now, "hh:mm AM/PM")
    End If
End Sub

Private Sub mnuLoad_Click()
On Error GoTo LoadFileErr
frmCalendar.Show
    
    ' This is where the data can be loaded for any date.  If the data is not
    ' in the ini file it tells the user that it is not there.  Otherwise the
    ' data is loaded into the appropriate fields
    If ReadINI(Format(dtSelected, "mm/dd/yyyy"), "Date", App.Path & "\Logtime.ini") = "" Then
        MsgBox "No Data for " & Format(dtSelected, "mm/dd/yyyy"), vbOKOnly + vbInformation, "No Data"
        Me.Caption = "Punch Clock " & Format(Now, "mm/dd/yyyy")
        
        dtSelected = Now
        ' Resets the data in the text fields
        For Index = 0 To 5
            txtTime(Index).Text = ""
        Next Index
        Me.Caption = "Punch Clock " & Format(dtSelected, "mm/dd/yyyy")
        txtTime(0).SetFocus
    Else
        x = ReadINI(Format(dtSelected, "mm/dd/yyyy"), "Date", App.Path & "\Logtime.ini")
        y = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(0).Caption, App.Path & "\Logtime.ini")
        z = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(1).Caption, App.Path & "\Logtime.ini")
        t = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(2).Caption, App.Path & "\Logtime.ini")
        v = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(3).Caption, App.Path & "\Logtime.ini")
        o = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(4).Caption, App.Path & "\Logtime.ini")
        p = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(5).Caption, App.Path & "\Logtime.ini")
        xx = ReadINI(Format(dtSelected, "mm/dd/yyyy"), "LunchOpt", App.Path & "\Logtime.ini")
        
        Me.Caption = "Punch Clock " & x
        txtTime(0).Text = y
        txtTime(1).Text = z
        txtTime(2).Text = t
        txtTime(3).Text = v
        txtTime(4).Text = o
        txtTime(5).Text = p
        If xx = "" Then
            xx = "0"
        Else
            chkLunch.Value = CInt(xx)
        End If
    End If
    Exit Sub
LoadFileErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
End Sub

Private Sub mnuPaste_Click()
    ' paste
    txtTime(intTextFieldLoc).Text = Clipboard.GetText
End Sub

Private Sub mnuSave_Click()
On Error GoTo SaveFileErr
    
    ' Validating only the first field for data when saving.  This is done
    ' so the user can save the initial data and come back to it later if
    ' necessary.
    If txtTime(0).Text = "" Then
        MsgBox "Please enter a " & Label1(x).Caption, vbOKOnly + vbCritical, "No " & Label1(x).Caption & " Entered"
        txtTime(0).SetFocus
        Exit Sub
    End If
    
    ' This how the data is saved to the ini file
    ' (Section, Key, Value, File location)
    WriteINI Format(dtSelected, "mm/dd/yyyy"), "Date", Format(Now, "mm/dd/yyyy"), App.Path & "\Logtime.ini"
    WriteINI Format(dtSelected, "mm/dd/yyyy"), Label1(0).Caption, txtTime(0).Text, App.Path & "\Logtime.ini"
    WriteINI Format(dtSelected, "mm/dd/yyyy"), Label1(1).Caption, txtTime(1).Text, App.Path & "\Logtime.ini"
    WriteINI Format(dtSelected, "mm/dd/yyyy"), Label1(2).Caption, txtTime(2).Text, App.Path & "\Logtime.ini"
    WriteINI Format(dtSelected, "mm/dd/yyyy"), Label1(3).Caption, txtTime(3).Text, App.Path & "\Logtime.ini"
    WriteINI Format(dtSelected, "mm/dd/yyyy"), Label1(4).Caption, txtTime(4).Text, App.Path & "\Logtime.ini"
    WriteINI Format(dtSelected, "mm/dd/yyyy"), Label1(5).Caption, txtTime(5).Text, App.Path & "\Logtime.ini"
    WriteINI Format(dtSelected, "mm/dd/yyyy"), "LunchOpt", chkLunch.Value, App.Path & "\Logtime.ini"
    Exit Sub
SaveFileErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
End Sub

Private Sub mnuTip_Click()
    ' Save whether or not this form should be displayed at startup
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", 1
    frmTip.Show
End Sub

Private Sub txtTime_GotFocus(Index As Integer)
    ' Allows for the user to tab through the fields and highlight all data
    ' without doing it manually
    SendKeys "{Home}+{End}"
    intTextFieldLoc = Index
End Sub

Private Sub txtTime_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbKeyTab
    End If
End Sub

Private Sub txtTime_LostFocus(Index As Integer)
    ' Another case where the code skips the Hour fields and formats the rest
    ' when lost focus
    Select Case Index
        Case 4
            Exit Sub
        Case 5
            Exit Sub
    End Select
    
    txtTime(Index).Text = Format(txtTime(Index).Text, "hh:mm AM/PM")
End Sub

Private Sub tbMAIN_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            ' Validating only the first field for data when saving.  This is done
            ' so the user can save the initial data and come back to it later if
            ' necessary.
            If txtTime(0).Text = "" Then
                MsgBox "Please enter a " & Label1(x).Caption, vbOKOnly + vbCritical, "No " & Label1(x).Caption & " Entered"
                txtTime(0).SetFocus
                Exit Sub
            End If
            
            Dim xQues As String, yTitle As String
            xQues = "Are you sure you want to save data for " & dtSelected & "?"
            yTitle = "Save data for " & dtSelected & "?"
            
            ' This how the data is saved to the ini file
            If MsgBox(xQues, vbYesNo + vbQuestion, yTitle) = vbYes Then
                ' (Section, Key, Value, File location)
                WriteINI Format(dtSelected, "mm/dd/yyyy"), "Date", Format(Now, "mm/dd/yyyy"), App.Path & "\Logtime.ini"
                WriteINI Format(dtSelected, "mm/dd/yyyy"), Label1(0).Caption, txtTime(0).Text, App.Path & "\Logtime.ini"
                WriteINI Format(dtSelected, "mm/dd/yyyy"), Label1(1).Caption, txtTime(1).Text, App.Path & "\Logtime.ini"
                WriteINI Format(dtSelected, "mm/dd/yyyy"), Label1(2).Caption, txtTime(2).Text, App.Path & "\Logtime.ini"
                WriteINI Format(dtSelected, "mm/dd/yyyy"), Label1(3).Caption, txtTime(3).Text, App.Path & "\Logtime.ini"
                WriteINI Format(dtSelected, "mm/dd/yyyy"), Label1(4).Caption, txtTime(4).Text, App.Path & "\Logtime.ini"
                WriteINI Format(dtSelected, "mm/dd/yyyy"), Label1(5).Caption, txtTime(5).Text, App.Path & "\Logtime.ini"
                WriteINI Format(dtSelected, "mm/dd/yyyy"), "LunchOpt", chkLunch.Value, App.Path & "\Logtime.ini"
            Else
                txtTime(0).SetFocus
                Exit Sub
            End If
        Case 2
            frmCalendar.Show
    
            ' This is where the data can be loaded for any date.  If the data is not
            ' in the ini file it tells the user that it is not there.  Otherwise the
            ' data is loaded into the appropriate fields
            If ReadINI(Format(dtSelected, "mm/dd/yyyy"), "Date", App.Path & "\Logtime.ini") = "" Then
                MsgBox "No Data for " & Format(dtSelected, "mm/dd/yyyy"), vbOKOnly + vbInformation, "No Data"
                Caption = "Punch Clock " & Format(Now, "mm/dd/yyyy")
                
                dtSelected = Now
                ' Resets the data in the text fields
                For Index = 0 To 5
                    frmTimeClocktxtTime.txtTime(Index).Text = ""
                Next Index
                Caption = "Punch Clock " & Format(dtSelected, "mm/dd/yyyy")
                txtTime(0).SetFocus
            Else
                x = ReadINI(Format(dtSelected, "mm/dd/yyyy"), "Date", App.Path & "\Logtime.ini")
                y = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(0).Caption, App.Path & "\Logtime.ini")
                z = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(1).Caption, App.Path & "\Logtime.ini")
                t = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(2).Caption, App.Path & "\Logtime.ini")
                v = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(3).Caption, App.Path & "\Logtime.ini")
                o = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(4).Caption, App.Path & "\Logtime.ini")
                p = ReadINI(Format(dtSelected, "mm/dd/yyyy"), Label1(5).Caption, App.Path & "\Logtime.ini")
                xx = ReadINI(Format(dtSelected, "mm/dd/yyyy"), "LunchOpt", App.Path & "\Logtime.ini")
        
                Caption = "Punch Clock " & x
                txtTime(0).Text = y
                txtTime(1).Text = z
                txtTime(2).Text = t
                txtTime(3).Text = v
                txtTime(4).Text = o
                txtTime(5).Text = p
                If xx = "" Then
                    xx = "0"
                Else
                    chkLunch.Value = CInt(xx)
                End If
            End If
        Case 4
            ' This is for setting the current time to text fields.  The code skips
            ' the Hours fields since those are for calculation purposes only
            If intTextFieldLoc > 3 Then
                Exit Sub
            Else
                txtTime(intTextFieldLoc).Text = Format(Now, "hh:mm AM/PM")
            End If
        Case 5
            frmCWW.Show
        Case 7
            'cut
            If txtTime(intTextFieldLoc).SelText = "" Then
                Exit Sub
            Else
                Clipboard.Clear
                Clipboard.SetText txtTime(intTextFieldLoc).SelText
                txtTime(intTextFieldLoc).SelText = ""
            End If
        Case 8
            'copy
            If txtTime(intTextFieldLoc).SelText = "" Then
                Exit Sub
            Else
                Clipboard.Clear
                Clipboard.SetText txtTime(intTextFieldLoc).SelText
            End If
        Case 9
            'paste
            txtTime(intTextFieldLoc).Text = Clipboard.GetText
    End Select
End Sub
