'Imports System.Runtime.InteropServices
'Imports System.Text

'Module OpenWindowGetter



'    Private Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr, ByVal lParam As Integer) As Boolean

'    <DllImport("USER32.DLL")>
'    Private Shared Function EnumWindows(ByVal enumFunc As EnumWindowsProc, ByVal lParam As Integer) As Boolean
'    End Function

'    <DllImport("USER32.DLL")>
'    Private Shared Function GetWindowText(ByVal hWnd As IntPtr, ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
'    End Function

'    <DllImport("USER32.DLL")>
'    Private Shared Function GetWindowTextLength(ByVal hWnd As IntPtr) As Integer
'    End Function

'    <DllImport("USER32.DLL")>
'    Private Shared Function IsWindowVisible(ByVal hWnd As IntPtr) As Boolean
'    End Function

'    <DllImport("USER32.DLL")>
'    Private Shared Function GetShellWindow() As IntPtr
'    End Function


'End Module










'Module mdMain

'    ' 這是 WINDOWPLACEMENT.showCmd 的常數

'    'Const SW_HIDE = 0
'    'Const SW_SHOWNORMAL = 1
'    'Const SW_NORMAL = 1
'    'Const SW_SHOWMINIMIZED = 2
'    'Const SW_SHOWMAXIMIZED = 3
'    'Const SW_MAXIMIZE = 3
'    'Const SW_SHOWNOACTIVATE = 4
'    'Const SW_SHOW = 5
'    'Const SW_MINIMIZE = 6
'    'Const SW_SHOWMINNOACTIVE = 7
'    'Const SW_SHOWNA = 8
'    'Const SW_RESTORE = 9
'    'Const SW_SHOWDEFAULT = 10
'    'Const SW_FORCEMINIMIZE = 11
'    'Const SW_MAX = 11

'    ' 這是 WINDOWPLACEMENT.rcNormalPosition 的集合變數, 不能用 VB 本身的 rectangel 來用, 因坐標不同, 及 參數是唯讀
'    Private Structure RECT
'        Public Left As Integer
'        Public Top As Integer
'        Public Right As Integer
'        Public Bottom As Integer
'    End Structure

'    ' 給 GetWindowPlacement 及 SetWindowPlacement 兩個 API 共用的参數
'    Private Structure WINDOWPLACEMENT
'        Public Length As Integer  ' 參數的長度, 必須要用 SizeOf 來設定, 才可使用
'        Public flags As Integer  ' 指令視窗的旗號 (1=最大視窗, 2=最小視窗, 4=線程同步 ... 一般開啟用 1, 縮小用 2)
'        Public showCmd As Integer  ' 視窗開啟後狀態的指令
'        Public ptMinPosition As Point  ' 視窗縮小後的位置
'        Public ptMaxPosition As Point  ' 視窗放大後的位置
'        Public rcNormalPosition As RECT  ' 視窗原本的位置及大小
'    End Structure

'    ' 兩個 API 副程式, 取得及設定視教狀態及位置及大小
'    Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Integer
'    Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Integer

'    ' 主程式
'    Public Sub Main()
'        Dim vCount As Integer = 0  ' 開啟的視窗數目
'        Dim wp As WINDOWPLACEMENT  ' API 用的參數

'        ' 若是取得的 ShowOut 旗號為 True, 表示剛才已縮小了視窗, 開始還原, 把視窗再次顯示
'        If GetSetting("Ranseco", "ShowDesktop", "ShowOut", False) Then
'            SaveSetting("Ranseco", "ShowDesktop", "ShowOut", False)  ' 重置旗號
'            ' 逐一打開視窗
'            For vCount = 1 To Val(GetSetting("Ranseco", "ShowDesktop", "End", 0))
'                wp.flags = 2  ' 設為放大
'                wp.showCmd = GetSetting("Ranseco", "ShowDesktop", "S" & CStr(vCount), 9)
'                ' 若取回的是 0 (隱藏視窗), 要改為 9 (重顯示)
'                If wp.showCmd = 0 Then wp.showCmd = 9
'                ' 取回視窗位置及大小的參數
'                wp.rcNormalPosition.Left = GetSetting("Ranseco", "ShowDesktop", "L" & CStr(vCount), 100)
'                wp.rcNormalPosition.Right = GetSetting("Ranseco", "ShowDesktop", "R" & CStr(vCount), 100)
'                wp.rcNormalPosition.Top = GetSetting("Ranseco", "ShowDesktop", "T" & CStr(vCount), 100)
'                wp.rcNormalPosition.Bottom = GetSetting("Ranseco", "ShowDesktop", "B" & CStr(vCount), 100)
'                wp.Length = System.Runtime.InteropServices.Marshal.SizeOf(wp)  ' 整理參數的大小, 存回 API (API 要求)
'                SetWindowPlacement(GetSetting("Ranseco", "ShowDesktop", "W" & CStr(vCount), 0), wp)  ' 執行指令
'            Next
'        Else
'            For Each p As Process In Process.GetProcesses
'                If p.MainWindowHandle <> 0 Then
'                    wp.Length = System.Runtime.InteropServices.Marshal.SizeOf(wp)
'                    GetWindowPlacement(p.MainWindowHandle, wp)
'                    'If (wp.showCmd <> 2) And (wp.rcNormalPosition.Left <> 0) And (p.MainWindowTitle <> "設定") Then
'                    If (wp.showCmd <> 2) And (wp.rcNormalPosition.Left <> 0) Then
'                        vCount += 1
'                        ' Debug.Print(p.MainWindowHandle.ToString & "---" & p.MainWindowTitle)  ' 用作除錯
'                        ' 把視窗的 句柄 (Handle) 儲存 及 記錄視窗位置及大小 ... (S 是視窗現狀, 一般是 1 = 正常, 3 = 最大)
'                        SaveSetting("Ranseco", "ShowDesktop", "W" & CStr(vCount), p.MainWindowHandle.ToString)
'                        SaveSetting("Ranseco", "ShowDesktop", "L" & CStr(vCount), wp.rcNormalPosition.Left)
'                        SaveSetting("Ranseco", "ShowDesktop", "R" & CStr(vCount), wp.rcNormalPosition.Right)
'                        SaveSetting("Ranseco", "ShowDesktop", "T" & CStr(vCount), wp.rcNormalPosition.Top)
'                        SaveSetting("Ranseco", "ShowDesktop", "B" & CStr(vCount), wp.rcNormalPosition.Bottom)
'                        SaveSetting("Ranseco", "ShowDesktop", "S" & CStr(vCount), wp.showCmd)
'                        wp.Length = System.Runtime.InteropServices.Marshal.SizeOf(wp)  ' 整理參數的大小 (API 要求)
'                        wp.flags = 1  ' 設為縮小
'                        wp.showCmd = 6  ' 設為縮小
'                        SetWindowPlacement(GetSetting("Ranseco", "ShowDesktop", "W" & CStr(vCount), 0), wp)  ' 執行指令
'                    End If
'                End If
'                ' 儲存本次最後處理視窗的最後編號
'                SaveSetting("Ranseco", "ShowDesktop", "End", vCount)
'                ' 若有視窗處理過, 設定旗號
'                If vCount > 0 Then SaveSetting("Ranseco", "ShowDesktop", "ShowOut", True)
'            Next
'        End If
'    End Sub
'End Module
