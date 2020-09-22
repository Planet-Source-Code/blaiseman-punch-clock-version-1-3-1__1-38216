VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCalendar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmCalendar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin MSACAL.Calendar Calendar1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2002
      Month           =   8
      Day             =   12
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()
    dtSelected = Calendar1.Value
    Me.Caption = "Calendar - Selected date: " & Format(dtSelected, "mm/dd/yyyy")
End Sub

Private Sub Calendar1_DblClick()
    ' Setting the public variable dtSelected to the date doubleclicked
    dtSelected = Calendar1.Value
    ' This is where the data can be loaded for any date.  If the data is not
    ' in the ini file it tells the user that it is not there.  Otherwise the
    ' data is loaded into the appropriate fields
    If ReadINI(Format(dtSelected, "mm/dd/yyyy"), "Date", App.Path & "\Logtime.ini") = "" Then
        MsgBox "No Data for " & Format(dtSelected, "mm/dd/yyyy"), vbOKOnly + vbInformation, "No Data"
        frmTimeClock.Caption = "Punch Clock " & Format(Now, "mm/dd/yyyy")
        
        dtSelected = Now
        ' Resets the data in the text fields
        For Index = 0 To 5
            frmTimeClock.txtTime(Index).Text = ""
        Next Index
        frmTimeClock.Caption = "Punch Clock " & Format(dtSelected, "mm/dd/yyyy")
        frmTimeClock.chkLunch.Value = 0
        frmTimeClock.txtTime(0).SetFocus
    Else
        x = ReadINI(Format(dtSelected, "mm/dd/yyyy"), "Date", App.Path & "\Logtime.ini")
        y = ReadINI(Format(dtSelected, "mm/dd/yyyy"), frmTimeClock.Label1(0).Caption, App.Path & "\Logtime.ini")
        z = ReadINI(Format(dtSelected, "mm/dd/yyyy"), frmTimeClock.Label1(1).Caption, App.Path & "\Logtime.ini")
        t = ReadINI(Format(dtSelected, "mm/dd/yyyy"), frmTimeClock.Label1(2).Caption, App.Path & "\Logtime.ini")
        v = ReadINI(Format(dtSelected, "mm/dd/yyyy"), frmTimeClock.Label1(3).Caption, App.Path & "\Logtime.ini")
        o = ReadINI(Format(dtSelected, "mm/dd/yyyy"), frmTimeClock.Label1(4).Caption, App.Path & "\Logtime.ini")
        p = ReadINI(Format(dtSelected, "mm/dd/yyyy"), frmTimeClock.Label1(5).Caption, App.Path & "\Logtime.ini")
        xx = ReadINI(Format(dtSelected, "mm/dd/yyyy"), "LunchOpt", App.Path & "\Logtime.ini")
        
        frmTimeClock.Caption = "Punch Clock " & x
        frmTimeClock.txtTime(0).Text = y
        frmTimeClock.txtTime(1).Text = z
        frmTimeClock.txtTime(2).Text = t
        frmTimeClock.txtTime(3).Text = v
        frmTimeClock.txtTime(4).Text = o
        frmTimeClock.txtTime(5).Text = p
        If xx = "" Then
            xx = "0"
        Else
            frmTimeClock.chkLunch.Value = CInt(xx)
        End If
        frmTimeClock.Caption = "Punch Clock " & Format(dtSelected, "mm/dd/yyyy")
    End If
    'Unload Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' Centering the form on load
    'CenterMe frmCalendar
    Me.Top = frmTimeClock.Top
    Me.Left = frmTimeClock.Left + 5010
    DockingStart Me
    
    ' Setting the Calendar's date to current day
    If Format(dtSelected, "mm/dd/yyyy") = "12/30/1899" Then
        dtSelected = Now
        Calendar1.Value = Now
    Else
        Calendar1.Value = dtSelected
    End If
    Me.Caption = "Calendar - Selected date: " & Format(dtSelected, "mm/dd/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'DockingTerminate Me
End Sub
