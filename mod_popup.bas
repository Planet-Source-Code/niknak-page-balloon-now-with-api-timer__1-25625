Attribute VB_Name = "mod_popup"
Option Explicit

'***************************************************************
'INFORMATION
'***************************************************************
    'SetTimer Me.hwnd, 0, 100, AddressOf TimerProc
    'KillTimer Me.hwnd, 0

'***************************************************************
'PUBLIC API DECLARATIONS
'***************************************************************
    'CREATES AN API TIMER
    Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    'KILLS AN API TIMER
    Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    'GETS WINDOW RECTANGLE
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    'CLOSES THE POPUP
    Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
    'USED TO KEEP FORM ONTOP
    Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'***************************************************************
'CONSTANTS
'***************************************************************
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
   
'***************************************************************
'PRIVATE TYPE DECLARATIONS
'***************************************************************
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type

'***************************************************************
'TIMER PROCEDURES
'***************************************************************
    'MOVE
    Public Sub moveProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
        Dim m_window As RECT
        GetWindowRect hwnd, m_window
        If Not busy Then
            If frm_main.chk_wrap Then
                If (m_window.Top * Screen.TwipsPerPixelY) > -((m_window.Bottom * Screen.TwipsPerPixelY) - (m_window.Top * Screen.TwipsPerPixelY)) Then
                    m_window.Top = m_window.Top - 2
                Else
                    m_window.Top = Screen.Height / Screen.TwipsPerPixelY
                End If
            Else
                If m_window.Top > 0 Then
                    m_window.Top = m_window.Top - 2
                Else
                    m_window.Top = 0
                End If
            End If
            If frm_main.chk_ontop Then
                SetWindowPos hwnd, HWND_TOPMOST, m_window.Left, m_window.Top, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE
            Else
                SetWindowPos hwnd, HWND_NOTOPMOST, m_window.Left, m_window.Top, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE
            End If
        End If
    End Sub

