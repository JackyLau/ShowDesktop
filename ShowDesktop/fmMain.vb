Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text

' ###################################################################################################
'   Auto Show Desktop ....... Version 1.0 ........ 2/3/2022
'   CopyRight Jacky Lau ... . Jackylau@yahoo.com
'
'   Use:
'   Execute the program, all open windows will be temporarily reduced to the bottom task bar
'   There is a small window near the middle of task bar, click it restore the previously window
'   The ICON of this program can be pinned to the task bar for easy access and use
'
'
'   使用:
'   執行程式, 所有已開啟的視窗會暫時縮至最底之工作列
'   在工作列上方中間位置, 有一小視窗, 再按一下便還原之前縮小的視窗
'   可以把此程式的 ICON, 釘至工作列上, 以方便取得使用
'
'
' ###################################################################################################

Public Class fmMain

    ' 這是 WINDOWPLACEMENT.showCmd 的常數, 留作參考

    'Const SW_HIDE = 0
    'Const SW_SHOWNORMAL = 1
    'Const SW_NORMAL = 1
    'Const SW_SHOWMINIMIZED = 2
    'Const SW_SHOWMAXIMIZED = 3
    'Const SW_MAXIMIZE = 3
    'Const SW_SHOWNOACTIVATE = 4
    'Const SW_SHOW = 5
    'Const SW_MINIMIZE = 6
    'Const SW_SHOWMINNOACTIVE = 7
    'Const SW_SHOWNA = 8
    'Const SW_RESTORE = 9
    'Const SW_SHOWDEFAULT = 10
    'Const SW_FORCEMINIMIZE = 11
    'Const SW_MAX = 11

    ' 給 GetWindowPlacement 及 SetWindowPlacement 兩個 API 共用的参數
    Private Structure WINDOWPLACEMENT
        Public Length As Integer  ' 參數的長度, 必須要用 SizeOf 來設定, 才可使用
        Public flags As Integer  ' 指令視窗的旗號 (1=最大視窗, 2=最小視窗, 4=線程同步 ... 一般開啟用 1, 縮小用 2)
        Public showCmd As Integer  ' 視窗開啟後狀態的指令
        Public ptMinPosition As Point  ' 視窗縮小後的位置
        Public ptMaxPosition As Point  ' 視窗放大後的位置
        Public rcNormalPosition As Rectangle  ' 視窗原本的位置及大小, 用系統中之 Rectangle 充作記錄 (注意: Left, Right, Top, Bottom 是不配合的)
    End Structure

    ' 要處理的視窗記錄
    Private Class cWinList
        Public vHandle As Integer  ' 儲存取回視窗的 句柄 (Handle)
        Public showCmd As Integer  ' 視窗原本之開啟狀態
        Public rcNormalPosition As Rectangle  ' 視窗原本的位置及大小
    End Class

    ' 線程 (Process) 的兩項參數
    Public Structure cProcess
        Public vHandle As Integer  ' 句柄 (Handle) 
        Public vTitle As String  ' 視窗名稱
    End Structure

    ' 兩個 API 副程式, 取得及設定視教狀態及位置及大小
    Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean

    Dim wp As WINDOWPLACEMENT  ' API 用的參數
    ReadOnly vWinListArray As New List(Of cWinList)  ' 要處理的視窗記錄陣列
    Public vProcessArray() As cProcess  ' 所有取回來 線程 (Process) 的 句柄 (Handle)

    ' 進入本程式, 處理縮小已開啟之視窗
    Private Sub fmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TopMost = True
        Me.Cursor = Cursors.NoMove2D
        Me.Size = New Drawing.Size(160, 80)
        Me.Location = New Point((My.Computer.Screen.WorkingArea.Width / 2) - (Me.Width / 2), My.Computer.Screen.WorkingArea.Height - Me.Height)

        ' 採用 齊整 線程 (足夠視窗) 方法, Enable 以下兩句, Remart Call pGetAllProcess() 的一句
        Dim cMoreIE As New csGetDesktopWin
        Call cMoreIE.pGetDesktopWin()
        'Call cMoreIE.GoGetOpenWin()

        ' 採用 簡單 線程 (自我編寫) 方法, Enable 以下一句, Remart 以上兩句
        'Call pGetAllProcess()

        Call pGetFileExplorer()  ' 再取檔案總管視窗
        Call pMinWindow()  ' 執行縮小視窗
    End Sub

    ' 按了本程式一下, 還原視窗, 程式結束
    Private Sub fmMain_Click(sender As Object, e As EventArgs) Handles Me.Click
        Call pMaxWindow()
        End
    End Sub

    ' 取得線程 (Process)
    Private Sub pGetAllProcess()
        Dim vCount As Integer  ' 執行計數
        Erase vProcessArray  ' 先清除原有資料

        ' 遂一取得正在執行中之程序
        For Each p As Process In Process.GetProcesses
            If p.MainWindowHandle <> 0 Then
                ReDim Preserve vProcessArray(vCount)  ' 重新增加陣列總數
                vProcessArray(vCount).vHandle = p.MainWindowHandle
                vProcessArray(vCount).vTitle = p.MainWindowTitle
                vCount += 1
            End If
        Next
    End Sub

    ' 取檔案總管視窗
    Private Sub pGetFileExplorer()
        Dim vCount As Integer
        Dim cShellWin As New SHDocVw.ShellWindows

        ' 取得之前已記錄, 準備處理的視窗數
        If Not IsNothing(vProcessArray) Then vCount = vProcessArray.Count

        ' 逐一取得視窗 (內裡名稱 InternetExplorer, 不是 IE, 是 "檔案總管" File Explorer)
        For Each OneShell As SHDocVw.InternetExplorer In cShellWin
            If OneShell.HWND <> 0 Then
                ReDim Preserve vProcessArray(vCount)
                vProcessArray(vCount).vHandle = OneShell.HWND
                vProcessArray(vCount).vTitle = OneShell.FullName
                vCount += 1
            End If
        Next
    End Sub

    ' 處理縮小已開啟之視窗
    Private Sub pMinWindow()
        ' 遂一取得已執行中之程序
        For Each p As cProcess In vProcessArray
            ' 找出有 Window 的程序

            wp.Length = System.Runtime.InteropServices.Marshal.SizeOf(wp)  ' 先整理一下參數的大小 (API 要求)
            GetWindowPlacement(p.vHandle, wp)  ' 取得該程序之 Window 的狀態
            If (wp.showCmd <> 2) And (wp.rcNormalPosition.Left <> 0) And (p.vTitle <> "設定") And (p.vTitle <> "Settings") Then

                Debug.Print(p.vHandle.ToString & "---" & p.vTitle)  ' 用作除錯

                ' 把視窗的 句柄 (Handle) 儲存 及 記錄視窗位置及大小 ... (S 是視窗現狀, 一般是 1 = 正常, 3 = 最大)
                Dim OneWinList As New cWinList With {
                        .vHandle = p.vHandle,
                        .showCmd = wp.showCmd,
                        .rcNormalPosition = wp.rcNormalPosition
                    }
                vWinListArray.Add(OneWinList)

                wp.flags = 1  ' 設為縮小
                wp.showCmd = 6  ' 設為縮小
                wp.Length = System.Runtime.InteropServices.Marshal.SizeOf(wp)  ' 整理參數的大小 (API 要求)
                SetWindowPlacement(OneWinList.vHandle, wp)  ' 執行指令, 把當時的視窗縮小
            End If
        Next
    End Sub

    ' 還原已縮小的視窗
    Private Sub pMaxWindow()
        Dim i As Integer
        ' 逐一打開視窗, 倒轉順序
        For i = (vWinListArray.Count - 1) To 0 Step -1
            wp.flags = 2  ' 設為放大

            ' 取回原有視窗狀態, 位置及大小的參數
            wp.showCmd = vWinListArray(i).showCmd
            If wp.showCmd = 0 Then wp.showCmd = 9  ' 若取回的是 0 (隱藏視窗), 要改為 9 (重顯示)
            wp.rcNormalPosition = vWinListArray(i).rcNormalPosition

            wp.Length = System.Runtime.InteropServices.Marshal.SizeOf(wp)  ' 整理參數的大小, 存回 API (API 要求)
            SetWindowPlacement(vWinListArray(i).vHandle, wp)  ' 執行指令, 把視窗還原
        Next
    End Sub
End Class

