Imports System.IO

Module LoadSetting
    Public SettingFile As String = "Settings.ini"
    Public machineListFile As String = "MachineList.ini"
    Public MachinePath, Destination, Source, Server As String
    Public ActionLog_path, ErrorLog_path, EmailLog_path As String
    Public actionLog, errorLog, emailLog As String

    Public SendEmail_Task As Boolean
    Public Send_mail_app_process As New System.Diagnostics.Process()
    Public Process_MailApp As New ProcessStartInfo
    Public Mail_App_Folder As String = ""
    Public Mail_App_Name As String = ""
    Public username, pass As String


    Sub LoadSettings()
        MachinePath = GetSettingItem(SettingFile, "machinePath") ' 1_CPUMEM
        Destination = GetSettingItem(SettingFile, "Destination") ' D:\data

        ActionLog_path = GetSettingItem(SettingFile, "ActionLog_path") & Format(Now, "yyyy") & "\" & Format(Now, "MM") & "\" 'Log\ActionLog\2024\09\
        actionLog = ActionLog_path & Format(Now, "dd") & ".txt"                                                             'Log\ActionLog\2024\09\14.txt


        ErrorLog_path = GetSettingItem(SettingFile, "ErrorLog_path") & Format(Now, "yyyy") & "\" & Format(Now, "MM") & "\"  'Log\ErrorLog\2024\
        errorLog = ErrorLog_path & Format(Now, "MM") & ".txt"                                                               'Log\ErrorLog\2024\09.txt

        EmailLog_path = GetSettingItem(SettingFile, "EmailLog_path") 'Log\Email\
        emailLog = EmailLog_path & "emailLog.txt"               '



        username = GetSettingItem(SettingFile, "ftp_user") 'temp_user
        pass = GetSettingItem(SettingFile, "ftp_pass")

        Call Create_Path(ActionLog_path)
        Call Create_Path(ErrorLog_path)
        Call Create_Path(EmailLog_path)


    End Sub

    Function GetSettingItem(ByVal file As String, ByVal searchItem As String) As String

        'Get setting value for variable
        Dim result = ""
        For Each line In System.IO.File.ReadLines(file)             'Pass=abc123
            If line.ToLower.Contains(searchItem.ToLower & "=") Then 'pass=
                result = line.Substring(searchItem.Length + 1)      '->abc123
                Exit For
            End If
        Next
        Return result

    End Function

    Sub Create_Path(ByVal path)
        Try
            If Directory.Exists(path) = False Then
                Directory.CreateDirectory(path)
            End If
        Catch ex As Exception
        End Try
    End Sub

End Module
