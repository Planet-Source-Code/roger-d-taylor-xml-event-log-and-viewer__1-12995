Attribute VB_Name = "modRoutines"
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Public Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)


Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public strEventLogLocation As String
Public strSource As String
Public strCategory As String
Public lngEvent As NTEventType
Public strDescription As String

Public Function GetUser() As String
    Dim lpUserID As String
    Dim nBuffer As Long
    Dim Ret As Long
    lpUserID = String(25, 0)
    nBuffer = 25
    Ret = GetUserName(lpUserID, nBuffer)


    If Ret Then
        GetUser$ = ClipNull(lpUserID)
    End If
    
End Function


Private Function ClipNull(InString As String) As String
    Dim intpos As Integer


    If Len(InString) Then
        intpos = InStr(InString, vbNullChar)


        If intpos > 0 Then
            ClipNull = Left(InString, intpos - 1)
        Else
            ClipNull = InString
        End If
    End If
End Function
