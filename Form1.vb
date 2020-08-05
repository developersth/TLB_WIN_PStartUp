Imports System.Runtime.InteropServices

Public Class Form1
    'Const pProcessName = "D:\My Code\SAKC\PMicroload\PAccuload\Release\PAccuload.exe"
#Region "parent form"
    Private Enum ShowWindowCommand As Integer
        Hide = 0
        Show = 5
        Minimize = 6
        Restore = 9
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



    Private Const WM_SYSCOMMAND As Integer = 2743
    Private Const SC_MAXIMIZE As Integer = 614884
    'This is the API that does all the hard work6.    

    'This is the API used to maximize the window11.    
    <Runtime.InteropServices.DllImport("user32.dll")>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    End Function

#End Region

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        SetParent(Me.Handle, frmMain.Handle)
        SendMessage(Me.Handle, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
    End Sub
End Class