Attribute VB_Name = "modMagnetic"
'"Magnetic Forms" Module.
'Copyright (C) 2001 Benjamin Wilger
'Benjamin@ActiveVB.de

'Modified by Adam Murray, 2002
' - Added support for a "master" form.  If any forms are docked to the master form
'   when it is moved, the docked forms move also.

Option Explicit

'Some API Declarations
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Declare Function SystemParametersInfo_Rect Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As RECT, ByVal fuWinIni As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10

Private Const WM_MOVING = &H216
Private Const WM_SIZING = &H214
Private Const WM_ENTERSIZEMOVE = &H231
Private Const WM_EXITSIZEMOVE = &H232

Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)

Private Const SPI_GETWORKAREA = 48

Private Const WMSZ_LEFT = 1
Private Const WMSZ_TOPLEFT = 4
Private Const WMSZ_BOTTOMLEFT = 7
Private Const WMSZ_TOP = 3
Private Const WMSZ_TOPRIGHT = 5

'User Declarations
'-----------------
Private Enum SnapFormMode
    Moving = 1
    Sizing = 2
End Enum

'We save the Infos in an UDT. That's easier to organize
Private Type DockingLog
    hwnd As Long
    oldProc As Long
End Type

Private m_hMasterWnd As Long

Private Logs() As DockingLog, LogCount As Integer, MaxLogs As Integer

Private MouseX As Long, MouseY As Long
Public SnappedX As Boolean, SnappedY As Boolean
Public Rects() As RECT

'Here, you can set the SnapWidth in Pixels. Ten's a good value.
Private Const SnapWidth = 10

'SubClassing is not very helpful while debugging your Code.
'If you need to step through your Code, set this Variable to False or
'you probably will crash!!!
Private Const DoSubClass As Boolean = True

'Deactivate Docking
Public Sub DockingTerminate(f As Form)
    Dim t As Integer, H As Long
    
    H = f.hwnd
    
    'delete entry as master form
    If m_hMasterWnd = H Then m_hMasterWnd = 0
    
    'Search Window
    For t = 0 To LogCount - 1
        If Logs(t).hwnd = H Then
            'Set back to Default WindowProc
            SetWindowLong H, GWL_WNDPROC, Logs(t).oldProc
            'Delete Window-Entry in Array
            For H = t To LogCount - 2
                Logs(H) = Logs(H + 1)
            Next H
            LogCount = LogCount - 1
            Exit For
        End If
    Next t
        
End Sub

'Activate Docking
Public Sub DockingStart(ByVal f As Form, Optional ByVal IsMaster As Boolean = False)
    Dim H As Long, t As Integer
    
    If Not DoSubClass Then Exit Sub
    
    'We redim only in 10 steps. This won't slow the Programm!
    If LogCount + 10 > MaxLogs Then
        MaxLogs = LogCount + 10
        ReDim Preserve Logs(MaxLogs)
    End If
    
    For t = 0 To LogCount - 1
        If Logs(t).hwnd = f.hwnd Then
            Debug.Print "Window-Docking already activated!"
            Exit Sub
        End If
    Next t

    H = f.hwnd
    Logs(LogCount).hwnd = H
    
    'Starting Subclassing and saving the old Window Procedure.
    Logs(LogCount).oldProc = SetWindowLong(H, GWL_WNDPROC, AddressOf WindowProc)

    'Set master status, if requested
    If IsMaster Then m_hMasterWnd = f.hwnd

    LogCount = LogCount + 1
    
End Sub

'This WindowProc will process all Messages coming from the
'Forms. The Messages we don't need will be redirected to the old Window Procedure
Public Function WindowProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim t As Integer ' Counter-Variable
    Dim oldProc As Long ' Address of original WindowProc
    Dim r As RECT, p As POINTAPI
    Dim runProc As Boolean
    Dim frm As Form
    runProc = True
    
    Dim rStartPos As RECT
    
    'Search Window in Array
    For t = 0 To LogCount - 1
        If Logs(t).hwnd = hwnd Then
            oldProc = Logs(t).oldProc
            Exit For
        End If
    Next t
    
    If oldProc = 0 Then Exit Function 'This would be not very good... ;-)
    
    If wMsg = WM_ENTERSIZEMOVE Then 'Windows tells us, that the User
                                    'begins to move or resize the Window.
        GetWindowRect hwnd, r
        GetCursorPos p
        MouseX = p.x - r.Left
        MouseY = p.y - r.Top
        
        GetFrmRects hwnd
        
    ElseIf wMsg = WM_SIZING Or wMsg = WM_MOVING Then 'While moving/sizing we're changing the Window Position/Size
                    
        'Get the rect info for the master window's current position (stored in twips)
        GetWindowRect hwnd, rStartPos
                    
        'Get the Rect-Structure from the Pointer located in lParam
        CopyMemory r, ByVal lParam, Len(r)
        
        'Change the Rect(see in DockFormRect)
        If wMsg = WM_SIZING Then
          DockFormRect hwnd, Sizing, r, wParam
        Else
          DockFormRect hwnd, Moving, r, wParam, MouseX, MouseY
        End If
        
        'Save it back.
        CopyMemory ByVal lParam, r, Len(r)
        
        'was this the master form we just moved?
        If hwnd = m_hMasterWnd Then
          
          Dim rTemp As RECT
          
          'examine all known docking-windows for their positions
          For t = 0 To LogCount - 1

            'but don't look at myself
            If Logs(t).hwnd <> hwnd Then
            
              'Get the window location of the candidate window
              GetWindowRect Logs(t).hwnd, rTemp
              
              'was this window docked to me in any way before i moved just now?
              If (rStartPos.Top = rTemp.Bottom) Or _
                 (rStartPos.Bottom = rTemp.Top) Or _
                 (rStartPos.Left = rTemp.Right) Or _
                 (rStartPos.Right = rTemp.Left) Then
                  
                'Calculate the delta for this window
                Dim nNewLeft As Long, nNewTop As Long
                nNewLeft = rTemp.Left + (r.Left - rStartPos.Left)
                nNewTop = rTemp.Top + (r.Top - rStartPos.Top)
                
                'Don't change the window's height and width...
                Dim nWidth As Long, nHeight As Long
                nWidth = rTemp.Right - rTemp.Left
                nHeight = rTemp.Bottom - rTemp.Top
                
                'update this Window's Position
                Call MoveWindow(Logs(t).hwnd, nNewLeft, nNewTop, nWidth, nHeight, 1)
                
              End If
            End If
          Next
          
        End If
        
        'Return a true Value(API uses 1 as True-Value)
        WindowProc = 1
        
        runProc = False 'Don't run OldWindowProc
    End If
    
    'Nachricht an originale Routine weiterleiten
    If runProc Then WindowProc = CallWindowProc(oldProc, hwnd, wMsg, wParam, lParam)
    
End Function

Private Function GetFrmRects(ByVal hwnd As Long)
  Dim frm     As Form
  Dim i       As Integer
  
  ReDim Rects(0 To 0)
  SystemParametersInfo_Rect SPI_GETWORKAREA, vbNull, Rects(0), 0
  
  i = 1
  
  For Each frm In Forms
    If frm.Visible And Not frm.hwnd = hwnd Then
      ReDim Preserve Rects(0 To i)
      GetWindowRect frm.hwnd, Rects(i)
      
      i = i + 1
    End If
  Next frm
End Function

'This is the heart of the Module. It automatically searches all
'visible Forms to dock on.
Private Sub DockFormRect(ByVal hwnd As Long, ByVal Mode As SnapFormMode, GivenRect As RECT, Optional SizingEdge As Long, Optional MouseX As Long, Optional MouseY As Long)
    Dim p As POINTAPI
    Dim i As Integer, diffX As Integer, diffY As Integer, diffWnd As Long
    Dim tmpRect As RECT, W As Integer, H As Integer, frmRect As RECT
    Dim XPos As Integer, YPos As Integer
    Dim tmpXPos As Integer, tmpYPos As Integer
    Dim tmpMouseX As Long, tmpMouseY As Long
    Dim FoundX As Boolean, FoundY As Boolean
    
    diffX = SnapWidth
    diffY = SnapWidth
    
    'Copy the original Rect.
    tmpRect = GivenRect
    
    frmRect = GivenRect
    
    'Do some calculations to correct the Window Position while Moving
    If Mode = Moving Then
        GetCursorPos p
        If SnappedX Then
            tmpMouseX = p.x - tmpRect.Left
            OffsetRect tmpRect, tmpMouseX - MouseX, 0
            OffsetRect GivenRect, tmpMouseX - MouseX, 0
        Else
            MouseX = p.x - tmpRect.Left
        End If
        If SnappedY Then
            tmpMouseY = p.y - tmpRect.Top
            OffsetRect tmpRect, 0, tmpMouseY - MouseY
            OffsetRect GivenRect, 0, tmpMouseY - MouseY
        Else
            MouseY = p.y - tmpRect.Top
        End If
    End If
    
    W = tmpRect.Right - tmpRect.Left
    H = tmpRect.Bottom - tmpRect.Top
    
    'that's the hard part!
    If Mode = Moving Then
        For i = 0 To UBound(Rects)
            If (tmpRect.Left >= (Rects(i).Left - SnapWidth) And _
                tmpRect.Left <= (Rects(i).Left + SnapWidth)) And _
                ((tmpRect.Top - SnapWidth) < Rects(i).Bottom And _
                (tmpRect.Bottom + SnapWidth) > Rects(i).Top) And _
                Abs(tmpRect.Left - Rects(i).Left) < diffX _
                Then
                
                GivenRect.Left = Rects(i).Left
                GivenRect.Right = GivenRect.Left + W
                
                diffX = Abs(tmpRect.Left - Rects(i).Left)
                
                FoundX = True
                
            ElseIf i > 0 And (tmpRect.Left >= (Rects(i).Right - SnapWidth) And _
                tmpRect.Left <= (Rects(i).Right + SnapWidth)) And _
                ((tmpRect.Top - SnapWidth) < Rects(i).Bottom And _
                (tmpRect.Bottom + SnapWidth) > Rects(i).Top) And _
                Abs(tmpRect.Left - Rects(i).Right) < diffX _
                Then
                
                GivenRect.Left = Rects(i).Right
                GivenRect.Right = GivenRect.Left + W
                
                diffX = Abs(tmpRect.Left - Rects(i).Right)
                
                FoundX = True
                
            ElseIf i > 0 And (tmpRect.Right >= (Rects(i).Left - SnapWidth) And _
                tmpRect.Right <= (Rects(i).Left + SnapWidth)) And _
                ((tmpRect.Top - SnapWidth) < Rects(i).Bottom And _
                (tmpRect.Bottom + SnapWidth) > Rects(i).Top) And _
                Abs(tmpRect.Right - Rects(i).Left) < diffX _
                Then
                
                GivenRect.Right = Rects(i).Left
                GivenRect.Left = GivenRect.Right - W
                
                diffX = Abs(tmpRect.Right - Rects(i).Left)
                
                FoundX = True
                
            ElseIf (tmpRect.Right >= (Rects(i).Right - SnapWidth) And _
                tmpRect.Right <= (Rects(i).Right + SnapWidth)) And _
                ((tmpRect.Top - SnapWidth) < Rects(i).Bottom And _
                (tmpRect.Bottom + SnapWidth) > Rects(i).Top) And _
                Abs(tmpRect.Right - Rects(i).Right) < diffX _
                Then
                
                GivenRect.Right = Rects(i).Right
                GivenRect.Left = GivenRect.Right - W
                
                diffX = Abs(tmpRect.Right - Rects(i).Right)
                
                FoundX = True
                
            End If
            
            'Y
            If (tmpRect.Top >= (Rects(i).Top - SnapWidth) And _
                tmpRect.Top <= (Rects(i).Top + SnapWidth)) And _
                ((tmpRect.Left - SnapWidth) < Rects(i).Right And _
                (tmpRect.Right + SnapWidth) > Rects(i).Left) And _
                Abs(tmpRect.Top - Rects(i).Top) < diffY _
                Then
                
                GivenRect.Top = Rects(i).Top
                GivenRect.Bottom = GivenRect.Top + H
                
                diffY = Abs(tmpRect.Top - Rects(i).Top)
                
                FoundY = True
                
            ElseIf i > 0 And (tmpRect.Top >= (Rects(i).Bottom - SnapWidth) And _
                tmpRect.Top <= (Rects(i).Bottom + SnapWidth)) And _
                ((tmpRect.Left - SnapWidth) < Rects(i).Right And _
                (tmpRect.Right + SnapWidth) > Rects(i).Left) And _
                Abs(tmpRect.Top - Rects(i).Bottom) < diffY _
                Then
                
                GivenRect.Top = Rects(i).Bottom
                GivenRect.Bottom = GivenRect.Top + H
                
                diffY = Abs(tmpRect.Top - Rects(i).Bottom)
                
                FoundY = True
                
            ElseIf i > 0 And (tmpRect.Bottom >= (Rects(i).Top - SnapWidth) And _
                tmpRect.Bottom <= (Rects(i).Top + SnapWidth)) And _
                ((tmpRect.Left - SnapWidth) < Rects(i).Right And _
                (tmpRect.Right + SnapWidth) > Rects(i).Left) And _
                Abs(tmpRect.Bottom - Rects(i).Top) < diffY _
                Then
                
                GivenRect.Bottom = Rects(i).Top
                GivenRect.Top = GivenRect.Bottom - H
                
                diffY = Abs(tmpRect.Bottom - Rects(i).Top)
                
                FoundY = True
                
            ElseIf (tmpRect.Bottom >= (Rects(i).Bottom - SnapWidth) And _
                tmpRect.Bottom <= (Rects(i).Bottom + SnapWidth)) And _
                ((tmpRect.Left - SnapWidth) < Rects(i).Right And _
                (tmpRect.Right + SnapWidth) > Rects(i).Left) And _
                Abs(tmpRect.Bottom - Rects(i).Bottom) < diffY _
                Then
                
                GivenRect.Bottom = Rects(i).Bottom
                GivenRect.Top = GivenRect.Bottom - H
                
                diffY = Abs(tmpRect.Bottom - Rects(i).Bottom)
                
                FoundY = True
                
            End If
        Next i
        
        'Save snapped state in Public Variable
        SnappedX = FoundX
        SnappedY = FoundY
        
    ElseIf Mode = Sizing Then
        If SizingEdge = WMSZ_LEFT Or SizingEdge = WMSZ_TOPLEFT Or SizingEdge = WMSZ_BOTTOMLEFT Then
            XPos = GivenRect.Left
        Else
            XPos = GivenRect.Right
        End If
        
        If SizingEdge = WMSZ_TOP Or SizingEdge = WMSZ_TOPLEFT Or SizingEdge = WMSZ_TOPRIGHT Then
            YPos = GivenRect.Top
        Else
            YPos = GivenRect.Bottom
        End If

        tmpXPos = XPos
        tmpYPos = YPos

        For i = 0 To UBound(Rects)
            If ((tmpXPos >= (Rects(i).Left - SnapWidth) And _
                tmpXPos <= (Rects(i).Left + SnapWidth)) And _
                ((tmpRect.Top - SnapWidth) < Rects(i).Bottom And _
                (tmpRect.Bottom + SnapWidth) > Rects(i).Top) And _
                Abs(tmpXPos - Rects(i).Left) < diffX) _
                Then

                XPos = Rects(i).Left
                
                diffX = Abs(tmpXPos - Rects(i).Left)
                
            ElseIf (tmpXPos >= (Rects(i).Right - SnapWidth) And _
                tmpXPos <= (Rects(i).Right + SnapWidth)) And _
                ((tmpRect.Top - SnapWidth) < Rects(i).Bottom And _
                (tmpRect.Bottom + SnapWidth) > Rects(i).Top) And _
                Abs(tmpXPos - Rects(i).Right) < diffX _
                Then
                
                XPos = Rects(i).Right
                
                diffX = Abs(tmpXPos - Rects(i).Right)
    
            End If
            
            'Y
            If (tmpYPos >= (Rects(i).Top - SnapWidth) And _
                tmpYPos <= (Rects(i).Top + SnapWidth)) And _
                ((tmpRect.Left - SnapWidth) < Rects(i).Right And _
                (tmpRect.Right + SnapWidth) > Rects(i).Left) And _
                Abs(tmpYPos - Rects(i).Top) < diffY _
                Then
                
                YPos = Rects(i).Top
                
                diffY = Abs(tmpYPos - Rects(i).Top)
                
            ElseIf (tmpYPos >= (Rects(i).Bottom - SnapWidth) And _
                tmpYPos <= (Rects(i).Bottom + SnapWidth)) And _
                ((tmpRect.Left - SnapWidth) < Rects(i).Right And _
                (tmpRect.Right + SnapWidth) > Rects(i).Left) And _
                Abs(tmpYPos - Rects(i).Bottom) < diffY _
                Then
                
                YPos = Rects(i).Bottom
                
                diffY = Abs(tmpYPos - Rects(i).Bottom)
            End If
        Next i

        If SizingEdge = WMSZ_LEFT Or SizingEdge = WMSZ_TOPLEFT Or SizingEdge = WMSZ_BOTTOMLEFT Then
            GivenRect.Left = XPos
        Else
            GivenRect.Right = XPos
        End If
    
        If SizingEdge = WMSZ_TOP Or SizingEdge = WMSZ_TOPLEFT Or SizingEdge = WMSZ_TOPRIGHT Then
            GivenRect.Top = YPos
        Else
            GivenRect.Bottom = YPos
        End If
    End If
    
End Sub
