Attribute VB_Name = "Module2"
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, _
                                        ByVal nIDEvent As Long) As Long
  
Private Declare Function SendMessageLongRef Lib "user32" _
        Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByRef lParam As Long) As Long
  
Private Declare Function FindWindow Lib "user32" _
        Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
                           
Private Declare Function FindWindowEx Lib "user32" _
        Alias "FindWindowExA" ( _
        ByVal hWnd1 As Long, _
        ByVal hWnd2 As Long, _
        ByVal lpsz1 As String, _
        ByVal lpsz2 As String) As Long
                           
Private Declare Function SetTimer Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long
                           
                           
Private m_ASC As Long
  
  
Sub inputbox_Password(El_Form As Form, Caracter As String)
      
    m_ASC = Asc(Caracter)
      
    Call SetTimer(El_Form.hwnd, &H5000&, 100, AddressOf TimerProc)
  
End Sub
  
  
Private Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, _
                                                                ByVal dwTime As Long)
           
    Dim Handle_InputBox As Long
      
    'Captura el handle del textBox del InputBox
    Handle_InputBox = FindWindowEx(FindWindow("#32770", App.Title), 0, "Edit", "")
                  
    'Le establece el PasswordChar
    Call SendMessageLongRef(Handle_InputBox, &HCC&, m_ASC, 0)
    'Finaliza el Timer
    Call KillTimer(hwnd, idEvent)
  
End Sub

