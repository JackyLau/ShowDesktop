Imports System.Runtime.InteropServices
Imports System.Text

Public Class csGetDesktopWin
    <DllImport("USER32.DLL")>
    Private Shared Function GetShellWindow() As IntPtr
    End Function

    <DllImport("USER32.DLL")>
    Private Shared Function GetWindowText(ByVal hWnd As IntPtr, ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function

    <DllImport("USER32.DLL")>
    Private Shared Function GetWindowTextLength(ByVal hWnd As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, <Out()> ByRef lpdwProcessId As UInt32) As UInt32
    End Function

    <DllImport("USER32.DLL")>
    Private Shared Function IsWindowVisible(ByVal hWnd As IntPtr) As Boolean
    End Function

    Private Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr, ByVal lParam As Integer) As Boolean

    <DllImport("USER32.DLL")>
    Private Shared Function EnumWindows(ByVal enumFunc As EnumWindowsProc, ByVal lParam As Integer) As Boolean
    End Function

    Private ReadOnly hShellWindow As IntPtr = GetShellWindow()
    Private ReadOnly dictWindows As New Dictionary(Of IntPtr, String)
    Private currentProcessID As Integer

    Public Function GetOpenWindowsFromPID(ByVal processID As Integer) As IDictionary(Of IntPtr, String)
        dictWindows.Clear()
        currentProcessID = processID
        EnumWindows(AddressOf enumWindowsInternal, 0)
        Return dictWindows
    End Function

    Private Function enumWindowsInternal(ByVal hWnd As IntPtr, ByVal lParam As Integer) As Boolean
        If (hWnd <> hShellWindow) Then
            Dim windowPid As UInt32
            If Not IsWindowVisible(hWnd) Then
                Return True
            End If
            Dim length As Integer = GetWindowTextLength(hWnd)
            If (length = 0) Then
                Return True
            End If
            GetWindowThreadProcessId(hWnd, windowPid)
            If (windowPid <> currentProcessID) Then
                Return True
            End If
            Dim stringBuilder As New StringBuilder(length)
            GetWindowText(hWnd, stringBuilder, (length + 1))
            dictWindows.Add(hWnd, stringBuilder.ToString)
        End If
        Return True
    End Function

    ' 另一取得視窗的方法 ... 可以取得副視窗, 例如 Edge 的所有已開啟的網頁視窗
    Public Sub pGetDesktopWin()
        Dim vCount As Integer = 0
        Erase fmMain.vProcessArray

        For Each poc In Process.GetProcesses
            If poc.MainWindowTitle.Length > 1 Then
                Dim windows As IDictionary(Of IntPtr, String) = GetOpenWindowsFromPID(poc.Id)
                For Each kvp As KeyValuePair(Of IntPtr, String) In windows
                    Try
                        ReDim Preserve fmMain.vProcessArray(vCount)
                        fmMain.vProcessArray(vCount).vHandle = kvp.Key
                        fmMain.vProcessArray(vCount).vTitle = kvp.Value
                        vCount += 1
                    Catch ex As Exception
                    End Try
                Next

            End If
        Next
    End Sub



    ' 再另一取得 Windows 的方法 ... 給 GoGetOpenWin 用
    Public Function GetOpenWindows() As IDictionary(Of IntPtr, String)
        Dim shellWindow As IntPtr = GetShellWindow()
        Dim windows As New Dictionary(Of IntPtr, String)()

        EnumWindows(Function(ByVal hWnd As IntPtr, ByVal lParam As Integer)
                        If hWnd = shellWindow Then Return True
                        If Not IsWindowVisible(hWnd) Then Return True
                        Dim length As Integer = GetWindowTextLength(hWnd)
                        If length = 0 Then Return True
                        Dim builder As StringBuilder
                        builder = New StringBuilder(length)
                        GetWindowText(hWnd, builder, length + 1)
                        windows(hWnd) = builder.ToString()
                        Return True
                    End Function, 0)
        Return windows
    End Function


    ' 再另一取得 Windows 的方法
    Public Sub GoGetOpenWin()
        Dim vCount As Integer = 0
        Erase fmMain.vProcessArray

        For Each OneWindow As KeyValuePair(Of IntPtr, String) In GetOpenWindows()
            ReDim Preserve fmMain.vProcessArray(vCount)
            fmMain.vProcessArray(vCount).vHandle = OneWindow.Key.ToString
            fmMain.vProcessArray(vCount).vTitle = OneWindow.Value
            vCount += 1
        Next
    End Sub

End Class

















