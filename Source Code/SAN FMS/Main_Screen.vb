Imports System.IO
Imports System.Net.Mail

Public Class Main_Screen

    Dim progresslabel As String = ""
    Dim attachmentpath As String = ""
    Dim filename As String = ""

    Dim attachmentsendfolder As String = ""
    Dim attachmentsentfolder As String = ""
    Dim attachmentdeletefolder As String = ""

    Dim shownminimizetip As Boolean = False

    Dim DateTimePicker1_Save As String = ""
    Dim DateTimePicker2_Save As String = ""

    Dim busyworking As Boolean = False

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message()
                If FullErrors_Checkbox.Checked = True Then
                    Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.ToString
                Else
                    Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.Message.ToString
                End If
                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                dir = Nothing
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ": " & ex.ToString)
                filewriter.WriteLine("")
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
                Label2.Text = "Error encountered in last action"

                '*********************
                'Send Mail Out

                If emailaddress.Text.Length > 0 And mailserver.Text.Length > 0 Then

                    Try
                        Dim obj As SmtpClient
                        If mailserverport.Text.Length > 0 Then
                            obj = New SmtpClient(mailserver.Text, mailserverport.Text)
                        Else
                            obj = New SmtpClient(mailserver.Text)
                        End If
                        Dim msg As MailMessage = New MailMessage

                        msg.Subject = "SAN FMS: Error Report (" & My.Computer.Name & ")"
                        Dim fromaddress As MailAddress = New MailAddress("unattended-mailbox@obe1.com.uct.ac.za", "SAN FMS Error Manager")
                        msg.From = fromaddress


                        For Each token As String In emailaddress.Text.Split(";")
                            msg.To.Add(token)
                        Next
                        obj.Send(msg)
                        obj = Nothing
                    Catch exc2 As Exception
                        MsgBox("An error occurred in the application's error handling routine (Sending Error Mail). The application will try to recover from this serious error." & vbCrLf & vbCrLf & exc2.ToString, MsgBoxStyle.Critical, "Critical Error Encountered")
                    End Try
                End If
            End If
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error." & vbCrLf & vbCrLf & exc.ToString, MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Email_Handler(ByVal identifier_msg As String)
        Try
            

            '*********************
            'Send Mail Out

            If emailaddress.Text.Length > 0 And mailserver.Text.Length > 0 Then

                Try
                    Dim obj As SmtpClient
                    If mailserverport.Text.Length > 0 Then
                        obj = New SmtpClient(mailserver.Text, mailserverport.Text)
                    Else
                        obj = New SmtpClient(mailserver.Text)
                    End If
                    Dim msg As MailMessage = New MailMessage

                    msg.Subject = "SAN FMS: Error Report (" & My.Computer.Name & ")"
                    Dim fromaddress As MailAddress = New MailAddress("unattended-mailbox@obe1.com.uct.ac.za", "SAN FMS Error Manager")
                    msg.From = fromaddress


                    For Each token As String In emailaddress.Text.Split(";")
                        msg.To.Add(token)
                    Next
                    obj.Send(msg)
                    obj = Nothing
                Catch exc2 As Exception
                    MsgBox("An error occurred in the application's email handling routine (Sending Error Mail). The application will try to recover from this serious error." & vbCrLf & vbCrLf & exc2.ToString, MsgBoxStyle.Critical, "Critical Error Encountered")
                End Try
            End If

        Catch exc As Exception
            MsgBox("An error occurred in the application's email handling routine. The application will try to recover from this serious error." & vbCrLf & vbCrLf & exc.ToString, MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub RunWorker()
        Try
            If busyworking = False Then
                busyworking = True
                Me.ControlBox = False

                Label2.Text = "Preparing to download Staff and Student list"
                progresslabel = ""
                ProgressBar1.Enabled = True
                ProgressBar1.Value = 0
                DateTimePicker1.Enabled = False
                DateTimePicker2.Enabled = False
                NumericUpDown1.Enabled = False
                CheckBox1.Enabled = False
                Label11.Enabled = False

                mailserver.Enabled = False
                mailserverport.Enabled = False
                emailaddress.Enabled = False
                sanbase.Enabled = False
                Button1.Enabled = False
                filename = ""
                attachmentpath = ""

                BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            Error_Handler(ex, "Run Worker")
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            shownminimizetip = False
            Control.CheckForIllegalCrossThreadCalls = False
            DateTimePicker1.Value = Now
            DateTimePicker2.Value = Now
            Me.Text = My.Application.Info.ProductName & " (" & Format(My.Application.Info.Version.Major, "0000") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ")"
            If My.Computer.FileSystem.DirectoryExists((Application.StartupPath & "\Temp").Replace("\\", "\")) = False Then
                My.Computer.FileSystem.CreateDirectory((Application.StartupPath & "\Temp").Replace("\\", "\"))
            End If
            loadSettings()
            Label2.Text = "Application loaded"
            Label7.Select()
        Catch ex As Exception
            Error_Handler(ex, "Form Load")
        End Try
    End Sub


    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            If sanbase.Text.Length > 0 And My.Computer.FileSystem.DirectoryExists(sanbase.Text) = True Then
                If emailaddress.Text.Length > 0 And mailserver.Text.Length > 0 Then
                    If My.Computer.Network.IsAvailable = True Then
                        If My.Computer.FileSystem.DirectoryExists((sanbase.Text & "\Staff").Replace("\\", "\")) = False Then
                            My.Computer.FileSystem.CreateDirectory((sanbase.Text & "\Staff").Replace("\\", "\"))
                        End If
                        If My.Computer.FileSystem.DirectoryExists((sanbase.Text & "\Students").Replace("\\", "\")) = False Then
                            My.Computer.FileSystem.CreateDirectory((sanbase.Text & "\Students").Replace("\\", "\"))
                        End If
                        Dim staffarray, studentarray As ArrayList
                        staffarray = New ArrayList
                        Dim binfo As DirectoryInfo
                        binfo = New DirectoryInfo((sanbase.Text & "\Staff").Replace("\\", "\"))
                        For Each dinfo As DirectoryInfo In binfo.GetDirectories()
                            staffarray.Add(dinfo.Name)
                            dinfo = Nothing
                        Next
                        binfo = Nothing
                        studentarray = New ArrayList
                        binfo = New DirectoryInfo((sanbase.Text & "\Students").Replace("\\", "\"))
                        For Each dinfo As DirectoryInfo In binfo.GetDirectories()
                            studentarray.Add(dinfo.Name)
                            dinfo = Nothing
                        Next
                        binfo = Nothing

                        Dim removestudents As Boolean = False
                        Dim removestaff As Boolean = False
                        '***************************
                        'Generate Staff Folders
                        'cn=00002119,ou=Staff,ou=com,ou=main,o=uct
                        progresslabel = "Downloading Staff List"
                        BackgroundWorker1.ReportProgress(0)
                        Dim lineread As String
                        Dim tempdownloadname As String = (Application.StartupPath & "\Temp\staff.txt").Replace("\\", "\")
                        My.Computer.Network.DownloadFile("http://www.commerce.uct.ac.za/Services/LDAP_Login/allcommercestaff_service.php?group=Staff_G", tempdownloadname, "", "", False, 100000, True)
                        progresslabel = "Staff List Downloaded"
                        BackgroundWorker1.ReportProgress(16)
                        progresslabel = "Creating Staff Folders"
                        BackgroundWorker1.ReportProgress(16)
                        Dim counter, newcounter As Integer
                        If My.Computer.FileSystem.FileExists(tempdownloadname) = True Then
                            Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(tempdownloadname)

                            counter = 0
                            newcounter = 0
                            While reader.Peek <> -1
                                lineread = reader.ReadLine.ToLower
                                If lineread.StartsWith("cn=") = True Then
                                    lineread = lineread.Remove(0, 3)
                                    lineread = lineread.Substring(0, lineread.IndexOf(",", 0))
                                    If IsNumeric(lineread) = True Then
                                        If lineread.Length = 8 Then
                                            counter = counter + 1
                                            removestaff = True
                                            lineread = lineread.ToUpper
                                            If staffarray.IndexOf(lineread) <> -1 Then
                                                staffarray.RemoveAt(staffarray.IndexOf(lineread))
                                            End If

                                            Try
                                                If My.Computer.FileSystem.DirectoryExists((sanbase.Text & "\Staff\" & lineread).Replace("\\", "\")) = False Then
                                                    My.Computer.FileSystem.CreateDirectory((sanbase.Text & "\Staff\" & lineread).Replace("\\", "\"))
                                                    newcounter = newcounter + 1
                                                End If
                                                Label2.Text = progresslabel & "(" & counter & "::" & newcounter & ")"
                                            Catch ex As Exception
                                                Error_Handler(ex, "Generating Folder: " & lineread)
                                            End Try
                                        End If
                                    End If
                                End If
                            End While
                            reader.Close()
                            reader = Nothing
                            My.Computer.FileSystem.DeleteFile(tempdownloadname)
                        End If
                        progresslabel = "Staff Folders Created"
                        BackgroundWorker1.ReportProgress(33)
                        progresslabel = "Removing Invalid Staff Folders"
                        BackgroundWorker1.ReportProgress(33)
                        counter = 0
                        newcounter = 0
                        If My.Computer.FileSystem.DirectoryExists((sanbase.Text & "\Removed").Replace("\\", "\")) = False Then
                            My.Computer.FileSystem.CreateDirectory((sanbase.Text & "\Removed").Replace("\\", "\"))
                        End If
                        If removestaff = True Then
                            For Each foldername As String In staffarray
                                counter = counter + 1

                                If My.Computer.FileSystem.DirectoryExists((sanbase.Text & "\Staff\" & foldername).Replace("\\", "\")) = True Then
                                    My.Computer.FileSystem.MoveDirectory((sanbase.Text & "\Staff\" & foldername).Replace("\\", "\"), (sanbase.Text & "\Removed\" & foldername & " " & Format(Now(), "yyyyMMddHHmmss")).Replace("\\", "\"))
                                    newcounter = newcounter + 1
                                End If
                                Label2.Text = progresslabel & "(" & counter & "::" & newcounter & ")"
                            Next
                        End If
                        progresslabel = "Invalid Staff Folders Removed"
                        BackgroundWorker1.ReportProgress(50)
                        '***************************
                        'Generate Student Folders
                        'cn=SMMTIM001,ou=Students,ou=com,ou=main,o=uct
                        progresslabel = "Downloading Student List"
                        BackgroundWorker1.ReportProgress(50)

                        tempdownloadname = (Application.StartupPath & "\Temp\students.txt").Replace("\\", "\")
                        My.Computer.Network.DownloadFile("http://www.commerce.uct.ac.za/Services/LDAP_Login/allcommercestudents_service.php?group=Students_G", tempdownloadname, "", "", False, 100000, True)
                        progresslabel = "Student List Downloaded"
                        BackgroundWorker1.ReportProgress(66)
                        progresslabel = "Creating Student Folders"
                        BackgroundWorker1.ReportProgress(66)
                        If My.Computer.FileSystem.FileExists(tempdownloadname) = True Then
                            Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(tempdownloadname)

                            counter = 0
                            newcounter = 0
                            While reader.Peek <> -1
                                lineread = reader.ReadLine.ToLower
                                If lineread.StartsWith("cn=") = True Then
                                    lineread = lineread.Remove(0, 3)
                                    lineread = lineread.Substring(0, lineread.IndexOf(",", 0))

                                    If lineread.Length = 9 Then
                                        If lineread.IndexOf("-") = -1 Then
                                            If IsNumeric(lineread.Substring(6, 3)) Then
                                                removestudents = True
                                                counter = counter + 1
                                                lineread = lineread.ToUpper
                                                If studentarray.IndexOf(lineread) <> -1 Then
                                                    studentarray.RemoveAt(studentarray.IndexOf(lineread))
                                                End If

                                                Try
                                                    If My.Computer.FileSystem.DirectoryExists((sanbase.Text & "\Students\" & lineread).Replace("\\", "\")) = False Then
                                                        My.Computer.FileSystem.CreateDirectory((sanbase.Text & "\Students\" & lineread).Replace("\\", "\"))
                                                        newcounter = newcounter + 1
                                                    End If
                                                    Label2.Text = progresslabel & "(" & counter & "::" & newcounter & ")"
                                                Catch ex As Exception
                                                    Error_Handler(ex, "Generating Folder: " & lineread)
                                                End Try
                                            End If
                                        End If
                                    End If
                                End If
                            End While
                            reader.Close()
                            reader = Nothing
                            My.Computer.FileSystem.DeleteFile(tempdownloadname)
                        End If
                        progresslabel = "Student Folders Created"
                        BackgroundWorker1.ReportProgress(83)
                        progresslabel = "Removing Invalid Student Folders"
                        BackgroundWorker1.ReportProgress(83)
                        counter = 0
                        newcounter = 0
                        If My.Computer.FileSystem.DirectoryExists((sanbase.Text & "\Removed").Replace("\\", "\")) = False Then
                            My.Computer.FileSystem.CreateDirectory((sanbase.Text & "\Removed").Replace("\\", "\"))
                        End If
                        If removestudents = True Then
                            For Each foldername As String In studentarray
                                counter = counter + 1

                                If My.Computer.FileSystem.DirectoryExists((sanbase.Text & "\Students\" & foldername).Replace("\\", "\")) = True Then
                                    My.Computer.FileSystem.MoveDirectory((sanbase.Text & "\Students\" & foldername).Replace("\\", "\"), (sanbase.Text & "\Removed\" & foldername & " " & Format(Now(), "yyyyMMddHHmmss")).Replace("\\", "\"))
                                    newcounter = newcounter + 1
                                End If
                                Label2.Text = progresslabel & "(" & counter & "::" & newcounter & ")"
                            Next
                        End If
                        progresslabel = "Invalid Student Folders Removed"
                        BackgroundWorker1.ReportProgress(100)


                        studentarray.Clear()
                        studentarray = Nothing
                        staffarray.Clear()
                        staffarray = Nothing






                    Else
                        e.Cancel = True
                        progresslabel = "No available network"
                        Email_Handler("No available network")
                        BackgroundWorker1.ReportProgress(100)
                    End If
                Else
                    e.Cancel = True
                    progresslabel = "No contact details specified"
                    BackgroundWorker1.ReportProgress(100)
                End If
            Else
                e.Cancel = True
                progresslabel = "SAN Base Folder not specified"
                Email_Handler("SAN Base Folder can not be accessed")
                BackgroundWorker1.ReportProgress(100)
            End If
        Catch ex As Exception
            Error_Handler(ex, "Background Worker: " & progresslabel)
            progresslabel = "Operation Failed: Error reported (" & progresslabel & ")"
        End Try
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            ProgressBar1.Value = e.ProgressPercentage
            Label2.Text = progresslabel
        Catch ex As Exception
            Error_Handler(ex, "Worker Progress Changed")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            ProgressBar1.Value = 100
            ProgressBar1.Enabled = False

            CheckBox1.Enabled = True
            If CheckBox1.Checked = True Then
                NumericUpDown1.Enabled = True
                DateTimePicker1.Enabled = False
                DateTimePicker2.Enabled = False
            Else
                NumericUpDown1.Enabled = False
                DateTimePicker1.Enabled = True
                DateTimePicker2.Enabled = True
            End If

            mailserver.Enabled = True
            mailserverport.Enabled = True
            emailaddress.Enabled = True

            sanbase.Enabled = True
            Button1.Enabled = True
            Label11.Enabled = True

            If e.Cancelled = True Then
                Label2.Text = "Failed to update SAN Folders: (" & progresslabel & ")"
            Else
                Label2.Text = "Operation Successfully Completed"
            End If

        Catch ex As Exception
            Error_Handler(ex, "Run Worker Completed")
        End Try
        Me.ControlBox = True
        busyworking = False
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Try
            Label2.Text = "About displayed"
            AboutBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display About Screen")
        End Try
    End Sub

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Try
            Label2.Text = "Help displayed"
            HelpBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display Help Screen")
        End Try
    End Sub



    Private Sub loadSettings()
        Try
            Label2.Text = "Loading application settings..."




            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            If My.Computer.FileSystem.FileExists(configfile) Then
                Dim reader As StreamReader = New StreamReader(configfile)
                Dim lineread As String
                Dim variablevalue As String
                While reader.Peek <> -1
                    lineread = reader.ReadLine
                    If lineread.IndexOf("=") <> -1 Then

                        variablevalue = lineread.Remove(0, lineread.IndexOf("=") + 1)

                        If lineread.StartsWith("sanbase=") Then
                            If My.Computer.FileSystem.DirectoryExists(variablevalue) = True Then
                                sanbase.Text = variablevalue
                                FolderBrowserDialog1.SelectedPath = variablevalue
                            End If
                        End If
                        If lineread.StartsWith("emailaddress=") Then
                            emailaddress.Text = variablevalue
                        End If
                        If lineread.StartsWith("mailserver=") Then
                            mailserver.Text = variablevalue
                        End If
                        If lineread.StartsWith("mailserverport=") Then
                            mailserverport.Text = variablevalue
                        End If
                        If lineread.StartsWith("DateTimePicker1=") Then
                            DateTimePicker1.Value = Date.Parse(variablevalue)
                            SaveSettings_Memory()
                        End If
                        If lineread.StartsWith("DateTimePicker2=") Then
                            DateTimePicker2.Value = Date.Parse(variablevalue)
                            SaveSettings_Memory()
                        End If
                        If lineread.StartsWith("NumericUpDown1=") Then
                            NumericUpDown1.Value = Integer.Parse(variablevalue)
                        End If
                        If lineread.StartsWith("CheckBox1=") Then
                            CheckBox1.Checked = Boolean.Parse(variablevalue)
                        End If
                        If lineread.StartsWith("FullErrors_Checkbox=") Then
                            FullErrors_Checkbox.Checked = Boolean.Parse(variablevalue)
                        End If

                    End If
                End While
                reader.Close()
                reader = Nothing
            End If


            'default values
            If emailaddress.Text.Length < 1 Then
                emailaddress.Text = "com-webmaster@uct.ac.za"
            End If
            If mailserver.Text.Length < 1 Then
                mailserver.Text = "obe1.com.uct.ac.za"
            End If
            If mailserverport.Text.Length < 1 Then
                mailserverport.Text = "25"
            End If



            Label2.Text = "Application Settings successfully loaded"
        Catch ex As Exception
            Error_Handler(ex, "Load Settings")
        End Try
    End Sub


    Private Sub SaveSettings()
        Try
            Label2.Text = "Saving application settings..."
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            Dim writer As StreamWriter = New StreamWriter(configfile, False)

            If sanbase.Text.Length > 0 Then
                writer.WriteLine("sanbase=" & sanbase.Text)
            End If
            If emailaddress.Text.Length > 0 Then
                writer.WriteLine("emailaddress=" & emailaddress.Text)
            End If
            If mailserver.Text.Length > 0 Then
                writer.WriteLine("mailserver=" & mailserver.Text)
            End If
            If mailserverport.Text.Length > 0 Then
                writer.WriteLine("mailserverport=" & mailserverport.Text)
            End If

            LoadSettings_Memory()

            writer.WriteLine("DateTimePicker1=" & DateTimePicker1.Value.ToString)
            writer.WriteLine("DateTimePicker2=" & DateTimePicker2.Value.ToString)

            writer.WriteLine("NumericUpDown1=" & NumericUpDown1.Value.ToString)
            writer.WriteLine("CheckBox1=" & CheckBox1.Checked.ToString)
            writer.WriteLine("FullErrors_Checkbox=" & FullErrors_Checkbox.Checked.ToString)
            

            writer.Flush()
            writer.Close()
            writer = Nothing

            Label2.Text = "Application Settings successfully saved"

        Catch ex As Exception
            Error_Handler(ex, "Save Settings")
        End Try
    End Sub


    Private Sub Main_Screen_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            If CheckBox1.Checked = True Then
                LoadSettings_Memory()
            End If
            SaveSettings()
        Catch ex As Exception
            Error_Handler(ex, "Closing Application")
        End Try
    End Sub


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Label7.Text = Format(Now, "HH:mm:ss")
            If Label7.Text = Label8.Text Or Label7.Text = Label1.Text Then
                If Label7.Text = Label1.Text And CheckBox1.Checked = True Then
                    'ignore because we're running on the scheduled timer
                Else
                    RunWorker()
                    If CheckBox1.Checked = True Then
                        DateTimePicker1.Value = DateTimePicker1.Value.AddMinutes(NumericUpDown1.Value)
                    End If
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "Timer Ticking")
        End Try
    End Sub


    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Try
            Label8.Text = Format(DateTimePicker1.Value, "HH:mm:ss")
        Catch ex As Exception
            Error_Handler(ex, "Change Scheduled Time")
        End Try
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        Try
            Label1.Text = Format(DateTimePicker2.Value, "HH:mm:ss")
        Catch ex As Exception
            Error_Handler(ex, "Change Scheduled Time")
        End Try
    End Sub


    Private Sub NotifyIcon1_BalloonTipClicked(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotifyIcon1.BalloonTipClicked
        Try
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            NotifyIcon1.Visible = False
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Click on NotifyIcon")
        End Try
    End Sub


    Private Sub NotifyIcon1_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseClick
        Try
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            NotifyIcon1.Visible = False
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Click on NotifyIcon")
        End Try
    End Sub


    Private Sub NotifyIcon1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        Try
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            NotifyIcon1.Visible = False
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Click on NotifyIcon")
        End Try
    End Sub

    Private Sub Main_Screen_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try
            If Me.WindowState = FormWindowState.Minimized Then
                Me.ShowInTaskbar = False
                NotifyIcon1.Visible = True
                If shownminimizetip = False Then
                    NotifyIcon1.ShowBalloonTip(1)
                    shownminimizetip = True
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "Change Window State")
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Try
            If CheckBox1.Checked = True Then
                SaveSettings_Memory()
                NumericUpDown1.Enabled = True
                DateTimePicker1.Enabled = False
                DateTimePicker2.Enabled = False
                DateTimePicker2.Value = Now.AddSeconds(-2)
                DateTimePicker1.Value = DateTimePicker2.Value
                DateTimePicker1.Value = DateTimePicker1.Value.AddMinutes(NumericUpDown1.Value)
            Else

                DateTimePicker1.Enabled = True

                DateTimePicker2.Enabled = True

                LoadSettings_Memory()

                NumericUpDown1.Enabled = False
            End If

        Catch ex As Exception
            Error_Handler(ex, "Enable Interval based timer")
        End Try
    End Sub

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        Try
            If CheckBox1.Checked = True Then
                ' DateTimePicker1.Value = DateTimePicker1.Value.AddMinutes(NumericUpDown1.Value)
                DateTimePicker1.Value = Now().AddMinutes(NumericUpDown1.Value)
                'SaveSettings_Memory()
            End If

        Catch ex As Exception
            Error_Handler(ex, "Increase interval")
        End Try
    End Sub

    Private Sub SaveSettings_Memory()
        Try
            DateTimePicker1_Save = DateTimePicker1.Value.ToString
            DateTimePicker2_Save = DateTimePicker2.Value.ToString
           
        Catch ex As Exception
            Error_Handler(ex, "SaveSettings_Memory")
        End Try
    End Sub

    Private Sub LoadSettings_Memory()
        Try
            If DateTimePicker1_Save.Length > 0 Then
                DateTimePicker1.Value = Date.Parse(DateTimePicker1_Save)
            End If
            If DateTimePicker2_Save.Length > 0 Then
                DateTimePicker2.Value = Date.Parse(DateTimePicker2_Save)
            End If


        Catch ex As Exception
            Error_Handler(ex, "LoadSettings_Memory")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                sanbase.Text = FolderBrowserDialog1.SelectedPath
            End If
        Catch ex As Exception
            Error_Handler(ex, "Select SAN Base Folder")
        End Try
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click
        RunWorker()
    End Sub
End Class
