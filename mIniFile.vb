Public Module IniFile

    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Int32

    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Int32, _
    ByVal lpFileName As String) As Int32


    Public Sub WriteIni(ByVal iniFileName As String, ByVal Section As String, ByVal ParamName As String, ByVal ParamVal As String)
        Dim Result As Integer

        Result = WritePrivateProfileString(Section, ParamName, ParamVal, iniFileName)

    End Sub

    Public Function ReadIni(ByVal IniFileName As String, ByVal Section As String, ByVal ParamName As String, ByVal ParamDefault As String) As String

        Dim ParamVal As String
        Dim LenParamVal As Long
        Dim vRet As String = ""
        Try
            ParamVal = Space$(1024)
            LenParamVal = GetPrivateProfileString(Section, ParamName, ParamDefault, ParamVal, Len(ParamVal), IniFileName)

            vRet = Left$(ParamVal, LenParamVal)
        Catch ex As Exception

        End Try
        Return vRet
    End Function

    Private Function GetCurrentApp() As String
        Return My.Application.Info.DirectoryPath
    End Function

    '#Region "Config Ducument Storage System"

    '    Public Function GetConfigIni() As Boolean
    '        Call GetInitialIni()

    '    End Function

    '    Private Sub GetInitialIni()
    '        Dim strFilepath As String
    '        strFilepath = GetCurrentApp() & ("\QSTNNRConfig.ini")

    '        P_DbConfig = ReadIni(strFilepath, "LOGON", "Database", "")
    '        P_UserPswConfig = ReadIni(strFilepath, "LOGON", "User_Psw", "")
    '        P_ReportDNSConfig = ReadIni(strFilepath, "LOGON", "Report_DNS", "")
    '        P_KeyStringConfig = ReadIni(strFilepath, "LOGON", "KeyString", "")

    '        TIMEOUT_DEACTIVE = ReadIni(strFilepath, "TimeOut", "DEACTIVE", "")
    '        TIMEOUT_AUTOLOGOUT = ReadIni(strFilepath, "TimeOut", "AUTOLOGOUT", "")
    '    End Sub

    '    Public Sub Write_USER(ByVal sUserName As String)
    '        'Dim strFilepath As String
    '        'strFilepath = GetCurrentApp() & ("\QuestionnaireExamConfig.ini")
    '        'Call writeIni(strFilepath, "USER", "NAME", sUserName)
    '    End Sub

    '#End Region

End Module
