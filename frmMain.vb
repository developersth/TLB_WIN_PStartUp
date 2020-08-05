Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Threading
Imports System.Collections
Imports System.Diagnostics
Imports System.Threading.Thread
Imports System.Linq

Public Class frmMain

    'Const pProcessName = "D:\My Code\Develop\StartUp\WindowsApplication1\PMicroload1.exe"
    Const C0Width = 0.1      '% ratio with datagridview.width
    Const C1Width = 0.63
    Const C2Width = 0.25

    Enum _State
        _Stop = 0
        _Starting = 1
        _Start = 2
        _Stopping = 3
    End Enum

    Structure _AppStartup
        Public _FileName As String
        Public _Args As Integer
        Public _Title As String
        Public _MainWindowHandle As IntPtr
        Public _Start As Integer
        Public _RECT As RECT
        Public _Width As Long
        Public _Height As Long
        Public _hwnd As Long
        Public _State As _State
        Public _Parent As Integer
        Public _IsParent As Boolean
        Public _ThrStartApp As Thread
        Public _ThrStopApp As Thread
        Public _ThrAutoStart As Thread
        Public _AutoStart As String
        Public _StartDate As Date
        Public _hwndParent As IntPtr
    End Structure

    'Dim pProcessName = Directory.GetCurrentDirectory() & "/PMicroload1.exe"
    Public IniPath As String = Directory.GetCurrentDirectory() & "/AppStartup.ini"
    Dim titleName As String
    Public totalApp As Integer
    Dim appProcess As New List(Of Process)
    Public appStartup() As _AppStartup
    Dim tmp_hWnd As IntPtr
    Dim tabPage() As TabPage
    Dim appIndex As Integer
    Dim thrRun As Boolean
    Dim idleTime As Integer    'default idle in minute
    Dim startSearch As Boolean = False
    Public p As New System.Diagnostics.Process

#Region "Thread"
    Dim thrApp As Thread
    Dim thrLockStart As Object = New Object
    Dim thrLockStop As Object = New Object
    Dim thrLockObj As Object = New Object
    Dim thrShutdown = False
    Private Sub StartThread()
        thrApp = New System.Threading.Thread(AddressOf RunProcess)
        thrApp.Name = "thrSearchActiveProcess"
        thrApp.Start()
    End Sub

    Private Sub RunProcess()
        'Thread.Sleep(10000)
        'Application.DoEvents()
        Dim sumAppStart As Integer
        startSearch = True
        While (thrShutdown = False)
            'Application.DoEvents()
            'Thread.Sleep(3000)
            If startSearch = True Then
                'SearchActiveProcess()
                ScanActiveProcess()
            End If
            Thread.Sleep(3000)
        End While
    End Sub

    Private Sub ThreadAutoStartApp(ByVal pIndex As Integer)
        appStartup(pIndex)._ThrAutoStart = New Thread(AddressOf StartApp)
        appStartup(pIndex)._ThrAutoStart.Name = "thrAutoStartApp" & pIndex
        appStartup(pIndex)._ThrAutoStart.Start(pIndex)
        Thread.Sleep(3000)
    End Sub

    Private Sub ThreadStartApp(ByVal pIndex As Integer)
        appStartup(pIndex)._ThrStartApp = New Thread(AddressOf StartApp)
        appStartup(pIndex)._ThrStartApp.Name = "thrStartApp" & pIndex
        appStartup(pIndex)._ThrStartApp.Start(pIndex)
    End Sub

    Private Sub SearchActiveProcess(ByVal pIndex As Integer)
        Thread.Sleep(1000)
        'appStartup(pIndex)._StartDate = Now
        While (thrShutdown = False)
            If SearchActiveProcessByName(pIndex) Then
                Exit While
            ElseIf (appStartup(pIndex)._StartDate - Now).TotalSeconds > 30 Then
                Exit While
            End If
            Thread.Sleep(1000)
        End While
        'appStartup(pIndex)._Thread.Abort()
    End Sub

    Sub StartApp(ByVal pIndex As Integer)
        StartSelectedProcess(pIndex)
        Thread.Sleep(3000)
    End Sub

    Private Sub ThreadStopApp(ByVal pIndex As Integer)
        appStartup(pIndex)._ThrStopApp = New Thread(AddressOf StopApp)
        appStartup(pIndex)._ThrStopApp.Name = "thrStopApp" & pIndex
        appStartup(pIndex)._ThrStopApp.Start(pIndex)
    End Sub

    Private Sub StopApp(ByVal pIndex As Integer)
        StopSelectedProcess(pIndex)
        Thread.Sleep(2000)
    End Sub
#End Region

    Sub ReadCurrentAppConfig(ByVal pIndex As Integer)
        appStartup(pIndex)._FileName = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", pIndex + 1, "NAME", "")
        appStartup(pIndex)._Args = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", pIndex + 1, "ARGUMENT", "")
    End Sub

    Private Sub StartUp()
        Try

            totalApp = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", 0, "TOTAL", "")
            titleName = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", 0, "TITLENAME", "")

            ReDim appStartup(totalApp - 1)
            ReDim tabPage(totalApp - 1)

            For i As Integer = 0 To totalApp - 1
                appStartup(i)._Width = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", 0, "WIDTH", "")
                appStartup(i)._Height = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", 0, "HIGH", "")
                appStartup(i)._FileName = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", i + 1, "NAME", "")
                appStartup(i)._Args = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", i + 1, "ARGUMENT", "")
                appStartup(i)._Title = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", i + 1, "TITLE", "")
                appStartup(i)._MainWindowHandle = New IntPtr(Convert.ToInt64(IIf(String.IsNullOrEmpty(ReadIni(GetCurrentDirectory() & "/AppStartup.ini", i + 1, "hWnd", "")), 0, _
                                                       ReadIni(GetCurrentDirectory() & "/AppStartup.ini", i + 1, "hWnd", ""))))
                appStartup(i)._Parent = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", i + 1, "PARENT", "")
                appStartup(i)._AutoStart = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", i + 1, "AUTO_START", "")
                tabPage(i) = New TabPage
                tabPage(i).Text = appStartup(i)._Title
            Next

            TabControl1.TabPages.Clear()
            TabControl1.TabPages.AddRange(tabPage)

            'save window handle
            For i = 0 To tabPage.Length - 1
                tabPage(i) = TabControl1.TabPages(i)
            Next

            Me.Text = titleName
            'SearchActiveProcess()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub InitialDatagridview()
        Dim i As Integer
        'dgvMain.Columns.Clear()
        'dgvMain.ColumnCount = 3
        'dgvMain.Columns(0).Name = "#"
        'dgvMain.Columns(1).Name = "Application Name"
        'dgvMain.Columns(2).Name = "Status"

        With dgvMain
            .Visible = False
            .RowCount = totalApp

            For i = 0 To .RowCount - 1
                .Rows(i).Height = 50
                .Rows(i).Cells(0).Value = i + 1
                .Rows(i).Cells(1).Value = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", i + 1, "TITLE", "").ToString()
                .Rows(i).Cells(2).Value = "STOP"
                .Rows(i).Cells(3).Value = IIf(appStartup(i)._AutoStart = "N", "N/A", "AUTO START")
                '.Rows(i).Cells(3).Value = 
                

            Next
            .Visible = True

            'Dim Column As New DataGridViewButtonCell

        End With
    End Sub

    Public Sub UpdateDatagridView()
        With dgvMain
            .Visible = False
            For i = 0 To .RowCount - 1
                .Rows(i).Cells(3).Value = IIf(appStartup(i)._AutoStart = "N", "N/A", "AUTO START")
            Next
            .Visible = True
        End With
    End Sub

    Private Function GetCurrentDirectory() As String
        Return Directory.GetCurrentDirectory 'My.Application.Info.DirectoryPath
    End Function

    Private Sub InitialProcess()
        Dim i As Integer

        For i = 0 To totalApp - 1
            Dim pProcess As New Process
            pProcess.StartInfo.FileName = appStartup(i)._FileName
            pProcess.StartInfo.Arguments = appStartup(i)._Args

            appProcess.Insert(i, pProcess)
        Next

        'SearchActiveProcess()
        StartThread()
    End Sub

    'Private Sub SetParentForm(ByVal pIndex As Integer)
    '    If TabControl1.TabPages(pIndex).InvokeRequired Then
    '        TabControl1.TabPages(pIndex).Invoke(New Action(Of Integer)(AddressOf SetParentForm), pIndex)
    '    Else
    '        appStartup(pIndex)._hwndParent = SetParent(appStartup(pIndex)._MainWindowHandle, tabPage(pIndex).Handle)
    '    End If
    'End Sub
    Delegate Sub SetParentFormDelegate(ByVal pIndex As Integer)
    Private Sub SetParentForm(ByVal pIndex As Integer)
        If TabControl1.TabPages(pIndex).InvokeRequired Then
            TabControl1.TabPages(pIndex).Invoke(New SetParentFormDelegate(AddressOf SetParentForm), pIndex)
            'SetParent(appStartup(pIndex)._MainWindowHandle, tabPage(pIndex).Handle)
        Else
            If (SetParent(appStartup(pIndex)._MainWindowHandle, tabPage(pIndex).Handle)).ToInt32 > 0 Then
                WriteIni(IniPath, pIndex + 1, "PARENT", 1)
                appStartup(pIndex)._IsParent = True
            End If
        End If
    End Sub

    Sub SearchActiveProcess()
        Dim pIndex As Integer
        Dim vSearch As Boolean = False
        Dim vProcessNameSearch As String = ""
        Dim vName() As String
        Try

            For pIndex = 0 To appProcess.Count - 1
                vSearch = False
                If appStartup(pIndex)._Parent = 1 Then GoTo Nextt
                If appStartup(pIndex)._FileName.IndexOf("\") >= 0 Then
                    'Dim s As String = mAppStartup(i)._FileName.Substring(mAppStartup(i)._FileName.IndexOf("\") + 1).Trim()
                    'vProcessNameSearch = s.Substring(0, s.IndexOf(".exe")).ToLower()
                    vName = appStartup(pIndex)._FileName.Split("\")
                    vProcessNameSearch = vName(vName.GetUpperBound(0))
                    vProcessNameSearch = vProcessNameSearch.Substring(0, vProcessNameSearch.IndexOf(".exe")).ToLower()
                Else
                    vProcessNameSearch = appStartup(pIndex)._FileName.Substring(0, appStartup(pIndex)._FileName.IndexOf(".exe")).Trim().ToLower()
                End If
                For Each pp As Process In Process.GetProcesses

                    If pp.ProcessName.ToLower().Contains(vProcessNameSearch.ToLower()) Then

                        'If mAppStartup(pIndex)._State = _State._Start Then
                        '    vSearch = True
                        '    Exit For
                        'End If

                        vSearch = True
                        If appStartup(pIndex)._Title <> pp.MainWindowTitle Then
                            SetWindowText(pp.MainWindowHandle, appStartup(pIndex)._Title)
                        End If
                        If appStartup(pIndex)._State = _State._Start Or appStartup(pIndex)._State = _State._Stopping And appStartup(pIndex)._IsParent Then Exit For
                        pp.WaitForInputIdle(500)
                        appStartup(pIndex)._MainWindowHandle = FindWindow(vbNullString, appStartup(pIndex)._Title)
                        If appStartup(pIndex)._MainWindowHandle <> 0 Then
                            'mAppStartup(i)._MainWindowHandle = pp.MainWindowHandle
                            WriteIni(IniPath, pIndex + 1, "WINDOW_PID", pp.Id)
                            WriteIni(IniPath, pIndex + 1, "hWnd", appStartup(pIndex)._MainWindowHandle)
                            'WriteIni(IniPath, i + 1, "PARENT", 1)
                            If Not appStartup(pIndex)._IsParent Then
                                ShowWindow(appStartup(pIndex)._MainWindowHandle, ShowWindowCommand.SW_HIDE)
                                ShowProcessWindow(pIndex)
                            End If
                        Else
                            appStartup(pIndex)._MainWindowHandle = ReadIni(IniPath, pIndex + 1, "hWnd", "")
                        End If
                        'mAppStartup(i)._hwmd = FindWindow(vbNullString, mAppStartup(i)._Title)
                        'mAppStartup(i)._MainWindowHandle = New IntPtr(Convert.ToInt64(mAppStartup(i)._hwmd))
                        If (appStartup(pIndex)._State = _State._Stop Or appStartup(pIndex)._State = _State._Starting) Then
                            appProcess.Item(pIndex) = pp
                            If dgvMain.Item(2, pIndex).Value.ToString.ToLower() = "start" Then
                                appStartup(pIndex)._MainWindowHandle = ReadIni(IniPath, pIndex + 1, "hWnd", "")
                                GetWindowRect(appStartup(pIndex)._MainWindowHandle, appStartup(pIndex)._RECT)
                                'SetParent(mAppStartup(i)._MainWindowHandle, mTabPage(i).Handle)
                                'ShowWindow(mAppStartup(i)._MainWindowHandle, ShowWindowCommand.SW_HIDE)
                                ShowProcessWindow(pIndex)
                            Else
                                'SetTextDatagrid(2, i, "Start")
                                GetWindowRect(appStartup(pIndex)._MainWindowHandle, appStartup(pIndex)._RECT)
                                'SetParent(mAppStartup(i)._MainWindowHandle, mTabPage(i).Handle)
                                'ShowWindow(mAppStartup(i)._MainWindowHandle, ShowWindowCommand.SW_SHOWNORMAL)
                                ShowProcessWindow(pIndex)
                                'WriteIni(IniPath, i + 1, "WINDOW_PID", pp.Id)
                                'WriteIni(IniPath, i + 1, "hWnd", mAppStartup(i)._MainWindowHandle)
                            End If
                            'GetWindowRect(mAppStartup(i)._MainWindowHandle, mAppStartup(i)._RECT)
                            'mAppStartup(i)._Start = 1
                            appStartup(pIndex)._State = _State._Start
                            SetTextDatagrid(2, pIndex, ShowStateDesc(appStartup(pIndex)._State))
                            Exit For
                        End If
                        'mAppStartup(i)._Start = 1
                        'Exit For
                    End If
                    Thread.Sleep(5)
                Next
                If Not vSearch Then
                    appStartup(pIndex)._State = _State._Stop
                    If dgvMain.Item(2, pIndex).Value.ToString.ToLower.IndexOf("start") >= 0 Or dgvMain.Item(2, pIndex).Value.ToString.ToLower.IndexOf("stopping") >= 0 Then
                        SetTextDatagrid(2, pIndex, ShowStateDesc(appStartup(pIndex)._State))
                    End If
                    If appStartup(pIndex)._AutoStart = "Y" And appStartup(pIndex)._IsParent = False Then
                        ThreadAutoStartApp(pIndex)
                        Thread.Sleep(1000)
                    End If
                End If
Nextt:
                'Thread.Sleep(300)
            Next pIndex
        Catch ex As Exception

        End Try
    End Sub

    Sub ScanActiveProcess()
        Dim pIndex As Integer
        Dim vSearch As Boolean = False
        Dim vProcessNameSearch As String = ""
        Dim vName() As String
        Try

            For pIndex = 0 To appProcess.Count - 1
                vSearch = False
                'If appStartup(pIndex)._IsParent = True Then GoTo Nextt
                If appStartup(pIndex)._FileName.IndexOf("\") >= 0 Then
                    vName = appStartup(pIndex)._FileName.Split("\")
                    vProcessNameSearch = vName(vName.GetUpperBound(0))
                    vProcessNameSearch = vProcessNameSearch.Substring(0, vProcessNameSearch.IndexOf(".exe")).ToLower()
                Else
                    vProcessNameSearch = appStartup(pIndex)._FileName.Substring(0, appStartup(pIndex)._FileName.IndexOf(".exe")).Trim().ToLower()
                End If
                'For Each pp As Process In Process.GetProcesses
                Dim vProcess As IEnumerable(Of Process) = (From p In Process.GetProcesses() _
                                    Where p.ProcessName.ToLower() = vProcessNameSearch.ToLower())
                'Select p)
                If vProcess.Count = 0 Then GoTo Stopp
                Dim pp As Process = vProcess.ElementAt(0)

                vSearch = True
                If appStartup(pIndex)._Title <> pp.MainWindowTitle Then
                    SetWindowText(pp.MainWindowHandle, appStartup(pIndex)._Title)
                End If
                If (appStartup(pIndex)._State = _State._Start And appStartup(pIndex)._IsParent) Or appStartup(pIndex)._State = _State._Stopping Then GoTo Stopp 'Exit For
                pp.WaitForInputIdle(300)
                'pp.Refresh()
                appStartup(pIndex)._MainWindowHandle = FindWindow(vbNullString, appStartup(pIndex)._Title)
                If appStartup(pIndex)._MainWindowHandle <> 0 Then
                    WriteIni(IniPath, pIndex + 1, "WINDOW_PID", pp.Id)
                    WriteIni(IniPath, pIndex + 1, "hWnd", appStartup(pIndex)._MainWindowHandle)
                    If Not appStartup(pIndex)._IsParent Then
                        ShowWindow(appStartup(pIndex)._MainWindowHandle, ShowWindowCommand.SW_HIDE)
                        ShowProcessWindow(pIndex)
                    End If
                Else
                    appStartup(pIndex)._MainWindowHandle = ReadIni(IniPath, pIndex + 1, "hWnd", "")
                End If

                If (appStartup(pIndex)._State = _State._Stop Or appStartup(pIndex)._State = _State._Starting) Then
                    appProcess.Item(pIndex) = pp
                    If dgvMain.Item(2, pIndex).Value.ToString.ToLower() = "start" Then
                        appStartup(pIndex)._MainWindowHandle = ReadIni(IniPath, pIndex + 1, "hWnd", "")
                        GetWindowRect(appStartup(pIndex)._MainWindowHandle, appStartup(pIndex)._RECT)
                        ShowProcessWindow(pIndex)
                    Else
                        'SetTextDatagrid(2, i, "Start")
                        GetWindowRect(appStartup(pIndex)._MainWindowHandle, appStartup(pIndex)._RECT)
                        ShowProcessWindow(pIndex)
                    End If
                    'GetWindowRect(mAppStartup(i)._MainWindowHandle, mAppStartup(i)._RECT)
                    'mAppStartup(i)._Start = 1
                    appStartup(pIndex)._State = _State._Start
                    SetTextDatagrid(2, pIndex, ShowStateDesc(appStartup(pIndex)._State))
                    'Exit For
                End If
                If appStartup(pIndex)._IsParent = False Then
                    appStartup(pIndex)._MainWindowHandle = ReadIni(IniPath, pIndex + 1, "hWnd", "")
                    GetWindowRect(appStartup(pIndex)._MainWindowHandle, appStartup(pIndex)._RECT)
                    ShowProcessWindow(pIndex)
                End If
                Thread.Sleep(5)
                'Next
Stopp:
                If Not vSearch Then
                    appStartup(pIndex)._State = _State._Stop
                    If dgvMain.Item(2, pIndex).Value.ToString.ToLower.IndexOf("start") >= 0 Or dgvMain.Item(2, pIndex).Value.ToString.ToLower.IndexOf("stopping") >= 0 Then
                        SetTextDatagrid(2, pIndex, ShowStateDesc(appStartup(pIndex)._State))
                    End If
                    If appStartup(pIndex)._AutoStart = "Y" And appStartup(pIndex)._IsParent = False Then
                        ThreadAutoStartApp(pIndex)
                        Thread.Sleep(1000)
                    End If
                End If
Nextt:
                'Thread.Sleep(300)
            Next pIndex
        Catch ex As Exception

        End Try
    End Sub

    Private Function SearchActiveProcessByName(ByVal pIndex As Integer)
        Dim vSearch As Boolean = False
        Dim vProcessNameSearch As String = ""
        Dim vName() As String
        Try

            If appStartup(pIndex)._FileName.IndexOf("\") >= 0 Then
                'Dim s As String = mAppStartup(pIndex)._FileName.Substring(0, mAppStartup(pIndex)._FileName.IndexOf("\")).Trim()
                'vProcessNameSearch = s.Substring(0, s.IndexOf(".exe")).ToLower()
                vName = appStartup(pIndex)._FileName.Split("\")
                vProcessNameSearch = vName(vName.GetUpperBound(0))
                vProcessNameSearch = vProcessNameSearch.Substring(0, vProcessNameSearch.IndexOf(".exe")).ToLower()
            Else
                vProcessNameSearch = appStartup(pIndex)._FileName.Substring(0, appStartup(pIndex)._FileName.IndexOf(".exe")).Trim().ToLower()
            End If

            For Each pp As Process In Process.GetProcesses
                'If pp.ProcessName.ToLower().Contains(vProcessNameSearch.ToLower()) Then
                If pp.ProcessName.ToLower() = vProcessNameSearch Then
                    vSearch = True
                    pp.WaitForInputIdle(300)
                    'pp.Refresh()
                    appStartup(pIndex)._MainWindowHandle = FindWindow(vbNullString, appStartup(pIndex)._Title)
                    If appStartup(pIndex)._MainWindowHandle <> 0 Then
                        'mAppStartup(i)._MainWindowHandle = pp.MainWindowHandle
                        WriteIni(IniPath, pIndex + 1, "WINDOW_PID", pp.Id)
                        WriteIni(IniPath, pIndex + 1, "hWnd", appStartup(pIndex)._MainWindowHandle)
                        WriteIni(IniPath, pIndex + 1, "START_DATE", appStartup(pIndex)._StartDate)
                        'WriteIni(IniPath, pIndex + 1, "PARENT", 1)
                        ShowWindow(appStartup(pIndex)._MainWindowHandle, ShowWindowCommand.SW_HIDE)
                        GetWindowRect(appStartup(pIndex)._MainWindowHandle, appStartup(pIndex)._RECT)
                        ShowProcessWindow(pIndex)
                    Else
                        WriteIni(IniPath, pIndex + 1, "PARENT", 0)
                    End If
                    Return vSearch
                End If
                'Thread.Sleep(500)
            Next
        Catch ex As Exception

        End Try
        Return vSearch
    End Function

    Private Sub SetTextDatagrid(ByVal pCol As Integer, ByVal pRow As Integer, ByVal pText As String)
        dgvMain.Item(pCol, pRow).Value = pText
    End Sub

    Private Sub StartSelectedProcess(ByVal pIndex As Integer)
        Application.DoEvents()
        If dgvMain.Item(2, pIndex).Value.ToString.ToLower() = "start" Then Exit Sub
        SyncLock thrLockStart
            Try
                If thrShutdown = True Then Exit Sub
                Application.DoEvents()
                If dgvMain.Item(2, pIndex).Value.ToString.ToLower() = "start" Then Exit Sub
                appStartup(pIndex)._StartDate = Now
                appStartup(pIndex)._State = _State._Starting
                ReadCurrentAppConfig(pIndex)
                SetTextDatagrid(2, pIndex, ShowStateDesc(appStartup(pIndex)._State))
                WriteIni(IniPath, pIndex + 1, "START_DATE", appStartup(pIndex)._StartDate)
                SetTextDatagrid(4, pIndex, appStartup(pIndex)._StartDate)
                If SearchActiveProcessByName(pIndex) Then
                    MessageBox.Show("Another instance of this app is running.", "Warning " & appStartup(pIndex)._Title)
                    Exit Sub
                End If
                idleTime = 0
                'ReadCurrentAppConfig(pIndex)
                'appStartup(pIndex)._State = _State._Starting
                'WriteIni(iniPath, pIndex + 1, "PARENT", 1)
                'SetTextDatagrid(2, pIndex, ShowStateDesc(appStartup(pIndex)._State))
                p = appProcess.Item(pIndex)

                p.StartInfo.UseShellExecute = False

                p.StartInfo.FileName = GetCurrentDirectory() & "\" & appStartup(pIndex)._FileName
                p.StartInfo.Arguments = appStartup(pIndex)._Args
                p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
                p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                p.Start()
                Thread.Sleep(500)
                appStartup(pIndex)._MainWindowHandle = p.MainWindowHandle
                ShowWindow(appStartup(pIndex)._MainWindowHandle, ShowWindowCommand.SW_SHOWMINIMIZED)
                'appStartup(pIndex)._Thread = New Thread(AddressOf SearchActiveProcess)
                'appStartup(pIndex)._Thread.Start(pIndex)
            Catch ex As Exception
                MsgBox(ex.Message & vbNewLine & appStartup(pIndex)._FileName, vbExclamation)
            End Try
        End SyncLock
    End Sub

    Private Sub ExitProgram()
        thrRun = True
        thrShutdown = True
        thrLockStart = False
        thrLockStop = False
        thrLockObj = False
        UnparentAllForm()
        p.Dispose()

        For Each pProcess As Process In appProcess
            pProcess.Dispose()
        Next
        appProcess.Clear()

    End Sub

    Private Sub StopSelectedProcess(ByVal pIndex As Integer)
        SyncLock thrLockStop
            If thrShutdown = True Then Exit Sub
            Try
                Application.DoEvents()
                idleTime = 0
                WriteIni(IniPath, pIndex + 1, "PARENT", 0)
                appStartup(pIndex)._State = _State._Stopping
                SetTextDatagrid(2, pIndex, ShowStateDesc(appStartup(pIndex)._State))
                appStartup(pIndex)._IsParent = False
                SendMessage(appStartup(pIndex)._MainWindowHandle, WM_SYSCOMMAND, WM_CLOSE, 0)
                'appStartup(pIndex)._State = _State._Stopping
                'SetTextDatagrid(2, pIndex, ShowStateDesc(appStartup(pIndex)._State))
                'dgvMain.Item(2, pIndex).Value = "Stopping..."
                'SearchActiveProcess()
                'StartThread()
            Catch ex As Exception
                MsgBox(ex.Message, vbCritical)
            End Try
        End SyncLock
    End Sub

    Private Sub KillSelectProcess(ByVal pIndex As Integer)
        Try
            appProcess(pIndex).Kill()
            Dim pProcess As New Process
            pProcess.StartInfo.FileName = appStartup(pIndex)._FileName
            pProcess.StartInfo.Arguments = appStartup(pIndex)._Args
            WriteIni(IniPath, pIndex + 1, "PARENT", 0)
            appProcess(pIndex) = pProcess
            appStartup(pIndex)._IsParent = False
        Catch ex As Exception
            WriteIni(IniPath, pIndex + 1, "PARENT", 0)
            MsgBox(ex.Message, vbCritical)
        End Try
    End Sub

    Private Sub ShowProcessWindow(ByVal pIndex As Integer)
        Dim vWidth As Long
        Dim vHeight As Long

        'Thread.Sleep(100)
        Try
            If appStartup(pIndex)._Height = 0 Then
                vHeight = tabPage(pIndex).Height
            Else
                vHeight = appStartup(pIndex)._Height
            End If

            If appStartup(pIndex)._Width = 0 Then
                vWidth = tabPage(pIndex).Width
            Else
                vWidth = appStartup(pIndex)._Width
            End If
            tmp_hWnd = appStartup(pIndex)._MainWindowHandle
            If Not tmp_hWnd.Equals(IntPtr.Zero) Then
                Application.DoEvents()
                ' use ShowWindow to change app window state (minimize and hide it).

                MoveWindow(tmp_hWnd, -1, -24, vWidth, vHeight + 20, True)
                'MoveWindow(tmp_hWnd, 0, -24, mAppStartup(pIndex)._RECT.right - mAppStartup(pIndex)._RECT.left, _
                'mAppStartup(pIndex)._RECT.top - mAppStartup(pIndex)._RECT.bottom, True)

                SetParentForm(pIndex)
                Thread.Sleep(300)
                'WriteIni(IniPath, pIndex + 1, "PARENT", 1)
                'appStartup(pIndex)._IsParent = True
                ShowWindow(tmp_hWnd, ShowWindowCommand.SW_SHOWNORMAL)

                Thread.Sleep(300)
                'mAppStartup(pIndex)._IsParent = 1

                'SetParent(tmp_hWnd, mTabPage(pIndex).Handle)
                'SetParent(tmp_hWnd, 0)  'unparent form
            Else
                ' no window handle?
                MessageBox.Show("Unable to get a window handle!")
                WriteIni(IniPath, pIndex + 1, "PARENT", 0)
                appStartup(pIndex)._IsParent = False
            End If
            SetTabpage(TabControl1, pIndex)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ResizeDatagridview()
        dgvMain.Width = GroupBox1.Width * 0.97
        dgvMain.Height = GroupBox1.Height * 0.9
        If dgvMain.Columns.Count > 0 Then
            With dgvMain
                .Columns(0).Width = .Width * C0Width
                .Columns(1).Width = .Width * C1Width
                .Columns(2).Width = .Width * C2Width
            End With
        End If
    End Sub

    Private Sub UnparentAllForm()
        Application.DoEvents()
        With dgvMain
            For i As Integer = 0 To .RowCount - 1
                If .Item(2, i).Value.ToString.ToLower() = "start" Then
                    MoveWindow(appStartup(i)._MainWindowHandle, appStartup(i)._RECT.left, appStartup(i)._RECT.top, _
                               appStartup(i)._RECT.right - appStartup(i)._RECT.left, appStartup(i)._RECT.bottom - appStartup(i)._RECT.top, True)
                    ShowWindow(appStartup(i)._MainWindowHandle, ShowWindowCommand.SW_HIDE)
                    Thread.Sleep(100)
                    SetParent(appStartup(i)._MainWindowHandle, 0)
                    WriteIni(IniPath, i + 1, "PARENT", 0)
                    'ShowWindow(mAppStartup(i)._MainWindowHandle, ShowWindowCommand.SW_HIDE)
                End If
            Next
        End With

    End Sub

    Private Sub UnparentForm(ByVal pIndex As Integer)
        Application.DoEvents()
        With dgvMain
            .Item(2, pIndex).Value = "Stop"
            SetParent(appStartup(pIndex)._MainWindowHandle, 0)
            ShowWindow(appStartup(pIndex)._MainWindowHandle, ShowWindowCommand.SW_HIDE)
            WriteIni(IniPath, pIndex + 1, "PARENT", 0)
        End With

    End Sub

    Private Function ShowStateDesc(ByVal pState As Integer)
        Select Case pState
            Case 0
                Return "STOP"
            Case 1
                Return "Starting"
            Case 2
                Return "START"
            Case 3
                Return "Stopping"
        End Select
        Return 0
    End Function

    Private Sub fMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ExitProgram()
    End Sub

    Private Sub fMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        StartUp()
        InitialDatagridview()
        ResizeDatagridview()
        InitialProcess()
        Application.DoEvents()
        Timer1.Start()
        'If titleName.ToLower.IndexOf("microload") >= 0 Then
        '    Me.Icon = My.Resources.startupMicroload
        'End If
        'If titleName.ToLower.IndexOf("card reader") >= 0 Then
        '    Me.Icon = My.Resources.startupCR
        'End If
    End Sub

    Private Sub fMain_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        With GroupBox1
            .Width = Me.Width / 4.3
            .Height = Me.Height
        End With
        ResizeDatagridview()
    End Sub

    Delegate Sub SetTabpagesDelegate(ByRef pTabControl As TabControl, ByVal pPageIndex As Integer)
    Sub SetTabpage(ByRef pTabControl As TabControl, ByVal pPageIndex As Integer)
        If pTabControl.InvokeRequired Then
            pTabControl.Invoke(New SetTabpagesDelegate(AddressOf SetTabpage), pTabControl, pPageIndex)
            'pTabControl.SelectedIndex = pPageIndex
        Else
            pTabControl.SelectedIndex = pPageIndex
        End If
    End Sub
    Delegate Sub GetTabpageHandleDelegate(ByRef pTabControl As TabControl, ByVal pPageIndex As Integer)
    Function GetTabpageHandle(ByRef pTabControl As TabControl, ByVal pPageIndex As Integer) As IntPtr
        If pTabControl.InvokeRequired Then
            pTabControl.Invoke(New GetTabpageHandleDelegate(AddressOf GetTabpageHandle), pTabControl, pPageIndex)
            'Return pTabControl.TabPages(pPageIndex).Handle
        Else
            Return pTabControl.TabPages(pPageIndex).Handle
        End If
    End Function

#Region "api window function"
    Private Enum ShowWindowCommand As Integer
        SW_HIDE = 0
        SW_SHOWNORMAL = 1
        SW_SHOWMINIMIZED = 2
        SW_SHOW = 5
        SW_MAXIMIZE = 3
        SW_MINIMIZE = 6
        SW_RESTORE = 9

    End Enum

    ' Set a specified window's show state 
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Shared Function ShowWindow(ByVal hwnd As IntPtr, ByVal nCmdShow As ShowWindowCommand) As Boolean
    End Function

    ' Determine whether the specified window handle identifies an existing window
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function IsWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    ' Determines whether a specified window is minimized.
    Private Declare Auto Function IsIconic Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean

    ' variable to save window handle
    Private calc_hWnd As IntPtr

    <DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function SetParent(ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")> _
    Private Shared Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, _
                                     ByVal uFlags As UInteger) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function GetWindowRect(ByVal hWnd As IntPtr, <Out()> ByRef lpRect As RECT) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <StructLayoutAttribute(LayoutKind.Sequential)> _
    Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

    Private Const WM_SYSCOMMAND As Integer = 274    '&H112
    Private Const SC_MAXIMIZE As Integer = 61488
    Private Const WM_CLOSE As Integer = &HF060& '&H10
    'This is the API that does all the hard work6.    

    'This is the API used to maximize the window11.    
    <Runtime.InteropServices.DllImport("user32.dll")>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    End Function

    <DllImport("user32.dll")> _
    Public Shared Function MoveWindow(ByVal hWnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    End Function

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
                            (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long,
                             ByVal lParam As Long) As Long

    <DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, _
                                                    ByRef lpdwProcessId As Integer) As Integer
    End Function

    Private Declare Auto Function GetWindowText Lib "user32" _
                                    (ByVal hWnd As System.IntPtr, _
                                    ByVal lpString As System.Text.StringBuilder, _
                                    ByVal cch As Integer) As Integer

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Shared Function SetWindowText(ByVal hwnd As IntPtr, ByVal lpString As String) As Boolean
    End Function
#End Region

#Region "ButtonTest"
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'p.StartInfo.FileName = pProcessName
        p.StartInfo.Arguments = 1

        p.StartInfo.UseShellExecute = False
        p.StartInfo.WindowStyle = ProcessWindowStyle.Normal
        p.StartInfo.CreateNoWindow = vbTrue
        p.Start()

        Thread.Sleep(3000)

        SetParent(Me.p.MainWindowHandle, Me.Handle)
        SendMessage(Me.p.MainWindowHandle, WM_SYSCOMMAND, SC_MAXIMIZE, 0)

        'Dim f As New Form1()

        'f.MdiParent = Me

        'f.Show()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim p As New System.Diagnostics.Process
        'p.StartInfo.FileName = pProcessName
        'p.CloseMainWindow()

        'p.Close()
        'p.Kill()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'p.StartInfo.FileName = pProcessName
        ' wait until app is in idle state
        p.WaitForInputIdle(-1)
        ' get app main window handle
        Dim tmp_hWnd As IntPtr = p.MainWindowHandle

        If Not tmp_hWnd.Equals(IntPtr.Zero) Then
            ' use ShowWindow to change app window state (minimize and hide it).
            ShowWindow(tmp_hWnd, ShowWindowCommand.SW_SHOW)
            ShowWindow(tmp_hWnd, ShowWindowCommand.SW_RESTORE)
            'ShowWindow(tmp_hWnd, ShowWindowCommand.Hide)
            ' save handle for later use.
            'calc_hWnd = tmp_hWnd
        Else
            ' no window handle?
            MessageBox.Show("Unable to get a window handle!")
        End If

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        SetParent(Me.p.MainWindowHandle, Me.Handle)
        SendMessage(Me.p.MainWindowHandle, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
        'SetParent(p.MainWindowHandle, Me.GroupBox1.Handle)

        'p.StartInfo.FileName = Path.Combine(Application.StartupPath, "PAccuload.exe")
        'SetParent(p.MainWindowHandle, Me.Handle)
        'SetParent(parentedprocess.MainWindowHandle,Me.Panel1.Handle) Try adding a panel to your form and use this line instead of the above line.43. 44.            
        'Now lets maximize the window of the process45.            
        'SendMessage(p.MainWindowHandle, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
    End Sub
    ' flags for ShowWindow, for more see, http://msdn.microsoft.com/en-us/library/windows/desktop/ms633548%28v=vs.85%29.aspx

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'Form1 = Me.ParentForm()


        Dim f As New Form1()

        f.MdiParent = Me
        f.Show()
    End Sub
#End Region

    Private Sub dgvMain_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMain.CellClick
        'If e.RowIndex = -1 Then Exit Sub
        'If e.ColumnIndex <> 1 Then Exit Sub
        'If dgvMain.Item(2, e.RowIndex).Value.ToString.ToLower() = "start" Then Exit Sub
        'StartSelectedProcess(e.RowIndex)
    End Sub

    Private Sub dgvMain_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMain.CellDoubleClick
        If e.RowIndex = -1 Or dgvMain.Item(2, e.RowIndex).Value.ToString.ToLower() = "stop" Then Exit Sub
        'ShowProcessWindow(e.RowIndex)
        SearchActiveProcessByName(e.RowIndex)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        'Dim pProcess As New Process
        'pProcess.StartInfo.FileName = GetCurrentApp() & "\PCRBay.exe"
        'pProcess.StartInfo.Arguments = 5
        'pProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
        'pProcess.Start()
        'pProcess.WaitForInputIdle()
        ''ShowWindow(pProcess.MainWindowHandle, ShowWindowCommand.Hide)
        'Thread.Sleep(5000)
        ''ShowWindow(pProcess.MainWindowHandle, ShowWindowCommand.Hide)
        'SetParent(pProcess.MainWindowHandle, mTabPage(0).Handle)
        ''Thread.Sleep(5000)
        ''SetParent(pProcess.MainWindowHandle, 0)     'unparent form
        Dim vWidth As Long
        Dim vHeight As Long


        If appStartup(0)._Height = 0 Then
            vHeight = tabPage(0).Height
        Else
            vHeight = appStartup(0)._Height
        End If

        If appStartup(0)._Width = 0 Then
            vWidth = tabPage(0).Width
        Else
            vWidth = appStartup(0)._Width
        End If

        tmp_hWnd = appStartup(0)._MainWindowHandle
        If Not tmp_hWnd.Equals(IntPtr.Zero) Then
            Application.DoEvents()
            ' use ShowWindow to change app window state (minimize and hide it).
            MoveWindow(tmp_hWnd, -1, -24, vWidth, vHeight + 20, True)
        End If
    End Sub

    Private Sub dgvMain_CellMouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvMain.CellMouseUp
        'mouse right click
        If e.Button = Windows.Forms.MouseButtons.Right And e.RowIndex > -1 Then
            dgvMain.ClearSelection()
            dgvMain.Rows(e.RowIndex).Selected = True
            ContextMenuStrip1.Show(Cursor.Position)
            appIndex = e.RowIndex
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        StartSelectedProcess(appIndex)
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        'StopSelectedProcess(appIndex)
        ThreadStopApp(appIndex)
    End Sub

    Private Sub KILLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KILLToolStripMenuItem.Click

    End Sub

    Private Sub dgvMain_CellMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvMain.CellMouseClick
        If e.Button = Windows.Forms.MouseButtons.Right And e.RowIndex > -1 Then Exit Sub
        If e.RowIndex > -1 Then
            TabControl1.SelectedIndex = e.RowIndex
            Exit Sub
        End If
        'ShowProcessWindow(e.RowIndex)
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem4.Click
       
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox1.ShowDialog()
    End Sub

    Private Sub StartAllToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StartAllToolStripMenuItem.Click
        Try
            For i As Integer = 0 To appStartup.Length - 1
                ThreadStartApp(i)
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub StopAllToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles StopAllToolStripMenuItem1.Click
        Try
            For i As Integer = 0 To appStartup.Length - 1
                ThreadStopApp(i)
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetTimeCloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetTimeCloseToolStripMenuItem.Click
        frmTimeClose.ShowDialog()
    End Sub

    Private Sub SettingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SettingToolStripMenuItem.Click

        frmSetting.InitialDatagridview(Me)
        frmSetting.ShowDialog()
        'KillSelectProcess(appIndex)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Stop_Program()
    End Sub
    Private Sub Stop_Program()
        Dim vTime As String
        Dim nTime = Format(Now, "HH:mm:ss")
        Dim nDate = Format(Now, "dd")
        Dim vCheckbox As String
        Dim vDate_Stop As String
        Dim vDate() As String
        Dim vCheck As Boolean = False
        vTime = ReadIni(Directory.GetCurrentDirectory() & "/AppStartup.ini", 0, "TIME_STOP", "").ToString()
        vCheckbox = ReadIni(Directory.GetCurrentDirectory() & "/AppStartup.ini", 0, "AUTO_STOP_ENABLE", "").ToString()
        vDate_Stop = ReadIni(Directory.GetCurrentDirectory() & "/AppStartup.ini", 0, "DATE_STOP", "").ToString()
        Try
            vDate = vDate_Stop.Split(",")
            For i As Integer = 0 To vDate.Length
                If nDate = vDate(i) Then
                    vCheck = True
                End If
            Next
        Catch ex As Exception

        End Try
        If nTime = vTime And vCheckbox = "Y" And vCheck = True Then
            Try
                'For i As Integer = 0 To appStartup.Length - 1
                '    ThreadStopApp(i)
                'Next

                totalApp = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", 0, "TOTAL", "")

                For i As Integer = 0 To totalApp - 1
                    appStartup(i)._AutoStart = ReadIni(GetCurrentDirectory() & "/AppStartup.ini", i + 1, "AUTO_START", "")
                    If appStartup(i)._AutoStart = "Y" Then
                        ThreadStopApp(i)
                    End If
                Next

            Catch ex As Exception

            End Try
        End If
    End Sub

End Class
