Imports System.IO
Imports System.Net

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call LoadSettings()
        Call Process_Files_From_FTP()

        If SendEmail_Task = True Then
            If File.Exists(emailLog) Then
                Call Write_log(actionLog, "Start sending Email.")
                Call Run_SendMailApp()
                Call Write_log(actionLog, "Complete.")
            End If
        End If


    End Sub

    Sub Run_SendMailApp()

        Process_MailApp.FileName = Mail_App_Name                    'Send_Email.exe
        Process_MailApp.WorkingDirectory = Mail_App_Folder          'D:\Send Email\
        Send_mail_app_process = Process.Start(Process_MailApp)      'Run mail app

    End Sub


    Function CreateFtpRequest(ByVal ftpUri As String, ByVal method As String) As FtpWebRequest

        Dim request As FtpWebRequest = CType(WebRequest.Create(ftpUri), FtpWebRequest)
        request.Method = method
        request.Credentials = New NetworkCredential(username, pass)
        Return request

    End Function

    Function GetFilePart(ByVal fileName As String, ByVal machineID As String) As String

        If fileName.Contains("_" & machineID & "_") AndAlso fileName.EndsWith(".csv") Then

            ' Tách lấy phần đầu tiên trước dấu gạch dưới (_)
            Return fileName.Split("_"c)(0)
        End If
        Return Nothing ' Trả về Nothing nếu không đúng định dạng

    End Function

    Sub Process_Files_From_FTP()

        For Each MachineLine As String In File.ReadLines(machineListFile) ' MachineList.ini
            Dim split_MachineLine() = MachineLine.Split(vbTab) ' FCA4    172.16.23.155
            Dim machineID = split_MachineLine(0) '-> FCA4
            Dim IpAddress = split_MachineLine(1) '-> 172.16.23.155
            Dim ftp_server = "ftp://" & IpAddress
            Dim FTP_FullPath = ftp_server & "/" & MachinePath.Replace("\", "/") ' ftp://172.16.23.89/1_CPUMEM/"

            If Check_ping(IpAddress) = True Then
                If Check_FTPPath_Exist(ftp_server) Then
                    For Each file In CheckFilesInFTPPath(FTP_FullPath, machineID)
                        If Not String.IsNullOrEmpty(file) AndAlso Check_LastModified(Path.Combine(FTP_FullPath, file)) Then
                            Dim destinationPath As String = Path.Combine(Destination, file.Split("_"c)(0), machineID, "RMCE")
                            Create_Path(destinationPath)

                            CopyFileFromFTP(FTP_FullPath, file, Path.Combine(destinationPath, file))
                            CompareAndDeleteFile(Path.Combine(FTP_FullPath, file), Path.Combine(destinationPath, file))
                        End If
                    Next
                End If
            End If
        Next

    End Sub

    'Sub Process_Files_From_FTP()

    '    For Each MachineLine As String In File.ReadLines(machineListFile) ' MachineList.ini

    '        Dim split_MachineLine() = MachineLine.Split(vbTab) ' FCA4    172.16.23.155
    '        Dim machineID = split_MachineLine(0) '-> FCA4
    '        Dim IpAddress = split_MachineLine(1) '-> 172.16.23.155
    '        Dim FTP_FullPath = "ftp://" & Path.Combine(IpAddress, MachinePath).Replace("\", "/") 'ftp: //172.16.23.89/1_CPUMEM/

    '        ' Duyệt file trong FTP_FullPath
    '        'Dim filesToProcess = filesToProcess
    '        For Each file In CheckFilesInFTPPath(FTP_FullPath, machineID)
    '            ' Lấy phần tử đầu tiên trước dấu "_"
    '            Dim filePart As String = file.Split("_"c)(0)

    '            ' Kiểm tra thời gian ghi cuối cùng của file
    '            If Check_LastModified(Path.Combine(FTP_FullPath, file)) Then
    '                ' Copy file từ FTP đến thư mục đích
    '                Dim destinationPath As String = Path.Combine(Destination, filePart, machineID, "RMCE") '
    '                Create_Path(destinationPath) ' Tạo thư mục đích nếu chưa tồn tại

    '                ' Copy file
    '                CopyFileFromFTP(FTP_FullPath, file, Path.Combine(destinationPath, file))
    '                CompareAndDeleteFile(Path.Combine(FTP_FullPath, file), Path.Combine(destinationPath, file))
    '            End If
    '        Next
    '    Next

    'End Sub


    Function Check_FTPPath_Exist(ByVal ftpPath As String) As Boolean

        Try
            ' Tạo yêu cầu FTP để liệt kê các thư mục
            Dim request = CreateFtpRequest(ftpPath, WebRequestMethods.Ftp.ListDirectory)
            Using response As FtpWebResponse = CType(request.GetResponse(), FtpWebResponse)
                Using reader As New StreamReader(response.GetResponseStream())
                    Dim folderExists As Boolean = False
                    Dim line As String = reader.ReadLine()
                    While Not String.IsNullOrEmpty(line)
                        If line.Contains(MachinePath) Then
                            folderExists = True
                            Exit While
                        End If
                    End While

                    If Not folderExists Then
                        Write_log(actionLog, "Thư mục " & MachinePath & " không tồn tại.")
                    End If

                    Return folderExists

                End Using
            End Using

        Catch ex As WebException
            Write_log(errorLog, "Lỗi: " & ex.Message)
            Return False
        Catch ex As Exception
            Write_log(errorLog, "Lỗi: " & ex.Message)
            Return False
        End Try

    End Function
    

    Function CheckFilesInFTPPath(ByVal ftpPath As String, ByVal machineID As String) As List(Of String)

        Dim fileList As New List(Of String)()
        Try
            Dim request = CreateFtpRequest(ftpPath, WebRequestMethods.Ftp.ListDirectory)
            Using response As FtpWebResponse = CType(request.GetResponse(), FtpWebResponse)

                'ListDirectory sẽ trả về phản hồi chứa thông tin chi tiết của file/folder
                'Ex: 2024/09/11 2:45 AFM-1562_FCA4_20240720_00053612.csv

                Using reader As New StreamReader(response.GetResponseStream())
                    Dim fileDetails = reader.ReadToEnd().Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                    For Each fileDetail In fileDetails
                        Dim fileName = fileDetail.Split(" "c).Last() 'AFM-1562_FCA4_20240720_00053612.csv

                        ' Kết hợp logic GetFilePart để kiểm tra và lấy phần đầu của tên file
                        If fileName.Contains("_" & machineID & "_") AndAlso fileName.EndsWith(".csv") Then
                            Dim filePart As String = fileName.Split("_"c)(0)
                            fileList.Add(filePart & "_" & fileName)
                        End If
                    Next

                End Using
            End Using

        Catch ex As WebException
            Write_log(errorLog, "Error checking FTP path: " & ex.Message)
        End Try

        Return fileList

    End Function


    
    '=========================================

    Function Check_LastModified(ByVal filePath As String) As Boolean

        Try
            Dim request = CreateFtpRequest(filePath, WebRequestMethods.Ftp.ListDirectory)
            Using response As FtpWebResponse = CType(request.GetResponse(), FtpWebResponse)
                Dim lastModified As DateTime = DateTime.Parse(response.LastModified.ToString())
                If DateTime.Now.Subtract(lastModified).TotalSeconds > 3 Then
                    Return True
                End If
            End Using
        Catch ex As WebException
            Write_log(errorLog, "Error checking file timestamp: " & ex.Message)
        End Try

        Return False

    End Function


    Sub CopyFileFromFTP(ByVal ftpPath As String, ByVal fileName As String, ByVal destinationPath As String)

        Try
            Dim request As FtpWebRequest = CreateFtpRequest(Path.Combine(ftpPath, fileName), WebRequestMethods.Ftp.DownloadFile)

            Using response As FtpWebResponse = CType(request.GetResponse(), FtpWebResponse)
                Using responseStream As Stream = response.GetResponseStream()
                    Using fileStream As New FileStream(destinationPath, FileMode.Create)
                        responseStream.CopyTo(fileStream)
                    End Using
                End Using
            End Using

            Write_log(actionLog, "File copied to: " & destinationPath)
        Catch ex As WebException
            Write_log(errorLog, "Error copying file: " & ex.Message)
            Call Write_log(emailLog, "  " & ex.Message)
        End Try

    End Sub


    Sub Create_Path(ByVal path As String)

        Try
            If Not Directory.Exists(path) Then
                Directory.CreateDirectory(path)
            End If
        Catch ex As Exception
            Write_log(errorLog, "Error creating directory: " & ex.Message)
        End Try

    End Sub


    Sub CompareAndDeleteFile(ByVal ftpFullPath As String, ByVal destinationPath As String)
        Try

            Dim destFileInfo As New FileInfo(destinationPath)
            Dim destFileSize As Long = destFileInfo.Length

            Dim ftpFileSize As Long = GetFileSizeFromFTP(ftpFullPath)

            If ftpFileSize > 0 AndAlso ftpFileSize = destFileSize Then
                DeleteFileFromFTP(ftpFullPath)
            Else

                Write_log(errorLog, "File sizes do not match. Retrying...") '
                Call Process_Files_From_FTP()

            End If
        Catch ex As Exception
            Write_Log(errorLog, "Error comparing files: " & ex.Message)
        End Try
    End Sub


    Function GetFileSizeFromFTP(ByVal ftpFullPath As String) As Long

        Try
            Dim request As FtpWebRequest = CType(WebRequest.Create(ftpFullPath), FtpWebRequest)
            request.Method = WebRequestMethods.Ftp.GetFileSize
            request.Credentials = New NetworkCredential(username, pass)

            Using response As FtpWebResponse = CType(request.GetResponse(), FtpWebResponse)
                Return response.ContentLength
            End Using
        Catch ex As WebException
            Write_log(errorLog, "Error getting file size from FTP: " & ex.Message)
            Return -1 ' Trả về -1 nếu có lỗi
        End Try

    End Function


    Sub DeleteFileFromFTP(ByVal ftpFullPath As String)

        Try
            Dim request As FtpWebRequest = CType(WebRequest.Create(ftpFullPath), FtpWebRequest)
            request.Method = WebRequestMethods.Ftp.DeleteFile
            request.Credentials = New NetworkCredential(username, pass)

            Using response As FtpWebResponse = CType(request.GetResponse(), FtpWebResponse)
                Write_log(actionLog, "File deleted from FTP: " & ftpFullPath)
            End Using
        Catch ex As WebException
            Write_log(errorLog, "Error deleting file from FTP: " & ex.Message)
        End Try

    End Sub


    Sub Write_log(ByVal filename As String, ByVal comment As String)
        'Write to error log or action log
        My.Computer.FileSystem.WriteAllText(filename, Format(Now(), "yyyy/MM/dd HH:mm:ss  ") & comment & vbNewLine, True)
    End Sub

    Function Check_ping(ByVal IPAddress) As Boolean

        Try
            ' CHECK THE CONECTION TO MACHINE
            If My.Computer.Network.Ping(IPAddress) Then
                ' Call Write_Log(actionLog, "  ")
                Check_ping = True
                Dim content = "  Connected " & IPAddress
                Call Write_log(actionLog, content)
            Else
                Dim content = "  Unable to connect to " & IPAddress
                Call Write_log(actionLog, content)
                Call Write_log(errorLog, content)
                Check_ping = False
            End If
        Catch ex As Exception
            Dim content = "  Unable to connect to " & IPAddress
            Call Write_log(actionLog, content)
            Call Write_log(errorLog, content)
            Check_ping = False
        End Try

    End Function

End Class
