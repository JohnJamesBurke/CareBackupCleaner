Option Strict On
Option Explicit On

Imports System.IO
Imports System.Configuration
Public Class CareBackupCleaner

    Private mstrBackupLocation As String = ""
    Private mintExpiryDays As Integer = 0
    Private mintPollingTimeHours As Integer = 0
    Private mintPollingTimeMins As Integer = 0
    Private mintPollingTimeSeconds As Integer = 0

    Private IntervalTimer As System.Threading.Timer
    Private mblnTimerStarted As Boolean = False

    Protected Overrides Sub OnStart(ByVal args() As String)

        ' Event log that the service is starting
        Dim myLog As New EventLog()
        If Not myLog.SourceExists("CareBackupCleaner") Then
            myLog.CreateEventSource("CareBackupCleaner", "Care Backup Cleaner Log")
        End If

        myLog.Source = "CareBackupCleaner"
        myLog.WriteEntry("Care Backup Cleaner Log", "Service Started on  " &
                    Date.Today.ToShortDateString & " " &
                    CStr(TimeOfDay),
                    EventLogEntryType.Information)



        Try

            ' Get the app settings 
            mstrBackupLocation = ConfigurationManager.AppSettings.Get("BackupLocation")

            Integer.TryParse(ConfigurationManager.AppSettings.Get("ExpiryDays"), mintExpiryDays)
            Integer.TryParse(ConfigurationManager.AppSettings.Get("PollingTimeHours"), mintPollingTimeHours)
            Integer.TryParse(ConfigurationManager.AppSettings.Get("PollingTimeMins"), mintPollingTimeMins)
            Integer.TryParse(ConfigurationManager.AppSettings.Get("PollingTimeSeconds"), mintPollingTimeSeconds)

            myLog.Source = "CareBackupCleaner"
            myLog.WriteEntry("Care Backup Cleaner Log", "Expiry Days: " & mintExpiryDays.ToString & vbCrLf &
                                                        "Hours: " & mintPollingTimeHours.ToString & vbCrLf &
                                                        "Minutes: " & mintPollingTimeMins.ToString & vbCrLf &
                                                        "Seconds: " & mintPollingTimeSeconds.ToString, EventLogEntryType.Information)


            ' Set the timer event
            If mintPollingTimeHours = 0 AndAlso mintPollingTimeMins = 0 AndAlso mintPollingTimeSeconds = 0 Then
                ' Do not start timer.
                ' Log event veiwer item "No settings specified"
                myLog = New EventLog()
                If Not myLog.SourceExists("CareBackupCleaner") Then
                    myLog.CreateEventSource("CareBackupCleaner", "Care Backup Cleaner Log")
                End If

                myLog.Source = "CareBackupCleaner"
                myLog.WriteEntry("Care Backup Cleaner Log", "No app timer settings available.  " &
                    Date.Today.ToShortDateString & " " &
                    CStr(TimeOfDay),
                    EventLogEntryType.Information)



            Else
                Dim tsInterval As TimeSpan = New TimeSpan(mintPollingTimeHours, mintPollingTimeMins, mintPollingTimeSeconds)
                IntervalTimer = New System.Threading.Timer(New System.Threading.TimerCallback(AddressOf IntervalTimer_Elapsed), Nothing, tsInterval, tsInterval)
                mblnTimerStarted = True
            End If


        Catch ex As Exception
            ' Event log that the service is starting
            myLog = New EventLog()
            If Not myLog.SourceExists("CareBackupCleaner") Then
                myLog.CreateEventSource("CareBackupCleaner", "Care Backup Cleaner Log")
            End If

            myLog.Source = "CareBackupCleaner"
            myLog.WriteEntry("Care Backup Cleaner Log", "An error occurred starting the service!" & vbCrLf & ex.Message.ToString() & vbCrLf & vbCrLf &
                    Date.Today.ToShortDateString & " " &
                    CStr(TimeOfDay),
                    EventLogEntryType.Error)

        End Try



    End Sub

    Protected Overrides Sub OnStop()

        ' Disable the timer
        If mblnTimerStarted = True Then
            IntervalTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
            IntervalTimer.Dispose()
            IntervalTimer = Nothing

        End If


        ' Event log that the service is starting
        Dim myLog As New EventLog()
        If Not myLog.SourceExists("CareBackupCleaner") Then
            myLog.CreateEventSource("CareBackupCleaner", "Care Backup Cleaner Log")
        End If

        myLog.Source = "CareBackupCleaner"
        myLog.WriteEntry("Care Backup Cleaner Log", "Service Stopped on  " &
                    Date.Today.ToShortDateString & " " &
                    CStr(TimeOfDay),
                    EventLogEntryType.Information)


    End Sub

    Private Sub IntervalTimer_Elapsed(ByVal state As Object)

        Try

            Dim intFilesTidied As Integer = 0

            Dim folderName As String = mstrBackupLocation
            If Strings.Right(folderName, 1) <> "\" Then folderName += "\"

            Dim di As New DirectoryInfo(folderName)

            ' Get a reference to each file in that directory.

            Dim fiArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo

            ' Loop through each file and check the last modified date.
            ' Then compare it with the expiry days and if older we remove the file.

            Dim dteExpiryDate As Date = Date.Now.AddDays((mintExpiryDays * -1))
            For Each fri In fiArr

                Dim dteLastModified As Date = fri.LastWriteTime
                If dteLastModified < dteExpiryDate Then

                    intFilesTidied += 1

                    ' Get the file name
                    Dim strFileName As String = fri.Name
                    Dim strFullFileName As String = fri.FullName

                    Dim myLog As New EventLog()
                    myLog.Source = "CareBackupCleaner"
                    myLog.WriteEntry("Care Backup Cleaner Log", "Cleaning up file:  " & strFileName & " on " &
                                        Date.Today.ToShortDateString & " " &
                                        CStr(TimeOfDay),
                                        EventLogEntryType.Information)

                    ' Delete the file
                    fri.Delete()

                    myLog.Source = "CareBackupCleaner"
                    myLog.WriteEntry("Care Backup Cleaner Log", "File:  " & strFileName & " removed on " &
                                        Date.Today.ToShortDateString & " " &
                                        CStr(TimeOfDay),
                                        EventLogEntryType.Information)

                End If

            Next fri

            ' Event log the number of files archived
            If intFilesTidied > 0 Then

                Dim myLog As New EventLog()
                myLog.Source = "CareBackupCleaner"
                myLog.WriteEntry("Care Backup Cleaner Log", intFilesTidied.ToString & " files tidied on " &
                                        Date.Today.ToShortDateString & " " &
                                        CStr(TimeOfDay),
                                        EventLogEntryType.Information)

            End If

        Catch ex As Exception
            Dim myLog As New EventLog()
            myLog.Source = "CareBackupCleaner"
            myLog.WriteEntry("Care Backup Cleaner Log", "An error occurred removing the file(s):" & vbCrLf & ex.Message.ToString & " on (" &
                                        Date.Today.ToShortDateString & " " &
                                        CStr(TimeOfDay) & ")",
                                        EventLogEntryType.Error)

        End Try

    End Sub

End Class
