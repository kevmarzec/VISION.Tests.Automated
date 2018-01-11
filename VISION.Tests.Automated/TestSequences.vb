Imports ServerLib

Public Class TestSequences


    Public Shared Sub launchVision(vispath As String)

        Dim tempName As String

        If System.IO.File.Exists(vispath + "\Vision.exe") Then
            tempName = "Vision"
        ElseIf System.IO.File.Exists(vispath + "\VISION.exe") Then
            tempName = "VISION"
        End If

        Form1.logIt("Test - Launch Vision - START")


        Dim p() As Process
        p = Process.GetProcessesByName("vision")

        If p.Count > 0 Then
            Form1.logIt("Vision is running")
            For i As Integer = 0 To p.Count - 1
                p(i).Kill()
            Next i

            Form1.logIt("Current Session terminated")
        Else
            Form1.logIt("Vision is not running.")

        End If





        Form1.logIt("Launching Vision")
        Process.Start(vispath + "\" + tempName)

        p = Nothing
        ''Time delay 20 seconds
        Threading.Thread.Sleep(20000)

        p = Process.GetProcessesByName("vision")
        If p.Count > 0 Then
            Form1.logIt("Successfully opened Vision instance")
            'Form1.logIt("Test - Launch Vision - PASS")
            Form1.logIt2("0002", "LAUNCH VISION", "PASS")
            Form1.logIt("Test - Launch Vision - PASS")
            Form1.passTest = Form1.passTest + 1

        Else
            'Form1.logIt("Test - Launch Vision - FAIL")
            'Form1.logItReport("Test - Launch Vision - FAIL")
            Form1.logIt("Test - Launch Vision - FAIL")
            Form1.logIt2("0002", "LAUNCH VISION", "FAIL")
            Form1.failTest = Form1.failTest + 1
        End If

        Form1.logIt("Test - Launch Vision - END")
        'Form1.logItReport("Test - Launch Vision - END")
        Form1.totalTests = Form1.totalTests + 1
    End Sub


    Public Shared Sub licCheck()

        Form1.logIt("Test - License Check - START")
        'Form1.logItReport("Test - License Check - START")
        Dim vis As New VisionApplicationComponent
        vis.FullyAutomatedMode = True

        If vis.IsLicensed = True Then
            'Form1.logIt("Test - License Check - PASS")
            'Form1.logItReport("Test - License Check - PASS")
            Form1.logIt("Test - License Check - PASS")
            Form1.logIt2("0001", "LICENSE CHECK", "PASS")
            Form1.passTest = Form1.passTest + 1
        Else
            'Form1.logIt("Test - License Check - FAIL")
            '
            Form1.logIt("Test - License Check - FAIL")
            Form1.logIt2("0001", "LICENSE CHECK", "FAIL")
            Form1.failTest = Form1.failTest + 1
        End If
        vis = Nothing
        Form1.logIt("Test - License Check - END")
        'Form1.logItReport("Test - License Check - END")
        Form1.totalTests = Form1.totalTests + 1

    End Sub

    Public Shared Sub openProject(ByVal filename As String)

        Form1.logIt("Test - Open Project - START")
        Dim vis As Object = CreateObject("Vision.Application")
        Dim project As IATIVisionProjectInterface = vis.GetCurrentProjectInterface

        If project.IsOpen = True Then
            project.Close()
            Form1.logIt("Closing active Project")
        End If
        Form1.logIt("Opening Project" + filename)
        project.Open(filename)

        If project.IsOpen = True Then
            If project.FileName = filename Then
                'Form1.logIt("Test - Open Project - PASS")
                'Form1.logItReport("Test - Open Project - PASS")
                Form1.logIt("Test - Open Project - PASS")
                Form1.logIt2("0003", "OPEN PROJECT", "PASS")
                Form1.passTest = Form1.passTest + 1
            Else
                'Form1.logIt("Test - Open Project - FAIL")
                'Form1.logItReport("Test - Open Project - FAIL")
                Form1.logIt("Test - Open Project - FAIL")
                Form1.logIt2("0003", "OPEN PROJECT", "FAIL")
                Form1.failTest = Form1.failTest + 1
            End If

        End If

        vis = Nothing
        project = Nothing
        Form1.logIt("Test - Open Project - END")
        Form1.totalTests = Form1.totalTests + 1
    End Sub

    Public Shared Sub projectOnline()

        Form1.logIt("Test - Project Online - START")
        Dim vis As Object = CreateObject("Vision.Application")
        Dim project As IATIVisionProjectInterface = vis.GetCurrentProjectInterface

        If project.IsOpen = True Then
            Form1.logIt("Setting Project Online")
            project.Online = True
        End If


        If project.Online = True Then

            'Form1.logIt("Test - Project Online - PASS")
            'Form1.logItReport("Test - Project Online - PASS")
            Form1.logIt("Test - Project Online - PASS")
            Form1.logIt2("0004", "SET PROJECT ONLINE", "PASS")
            Form1.passTest = Form1.passTest + 1
        Else
            Form1.logIt("Test - Project Online - FAIL")
            Form1.logIt2("0004", "SET PROJECT ONLINE", "FAIL")
            'Form1.logItReport("Test - Project Online - FAIL")
            Form1.failTest = Form1.failTest + 1
        End If



        vis = Nothing
        project = Nothing
        Form1.logIt("Test - Project Online - END")
        Form1.totalTests = Form1.totalTests + 1
    End Sub

    Public Shared Sub projectOffline()

        Form1.logIt("Test - Project Offline - START")
        Dim vis As Object = CreateObject("Vision.Application")
        Dim project As IATIVisionProjectInterface = vis.GetCurrentProjectInterface

        If project.IsOpen = True Then
            Form1.logIt("Setting Project Offline")
            project.Online = False
        End If


        If project.Online = False Then
            Form1.logIt("Test - Project Offline - PASS")
            Form1.logIt2("0005", "SET PROJECT OFFLINE", "PASS")
            'Form1.logItReport("Test - Project Offline - PASS")
            Form1.passTest = Form1.passTest + 1
        Else
            Form1.logIt("Test - Project Offline - FAIL")
            Form1.logIt2("0005", "SET PROJECT OFFLINE", "FAIL")
            'Form1.logItReport("Test - Project Offline - FAIL")
            Form1.failTest = Form1.failTest + 1
        End If



        vis = Nothing
        project = Nothing
        Form1.logIt("Test - Project Offline - END")
        Form1.totalTests = Form1.totalTests + 1
    End Sub

    Public Shared Sub pcmDownload()

        Form1.logIt("Test - Download - START")
        Dim vis As Object = CreateObject("Vision.Application")
        Dim project As IATIVisionProjectInterface = vis.GetCurrentProjectInterface
        Dim pcm As IATIVisionDeviceInterface = project.FindDevice("PCM")
        Dim stat As Boolean



        Threading.Thread.Sleep(2000)
        Do

        Loop Until project.Online = True

        If project.Online = True Then
            Form1.logIt("Start Calibration Download")
            stat = pcm.DownloadActiveStrategy()
        End If

        If stat = True Then
            Form1.logIt("Test - Download - PASS")
            Form1.logIt2("0006", "CAL DOWNLOAD", "PASS")
            'Form1.logItReport("Test - Download - PASS")
            Form1.passTest = Form1.passTest + 1
        Else
            Form1.logIt("Test - Download - FAIL")
            Form1.logIt2("0006", "CAL DOWNLOAD", "FAIL")
            'Form1.logItReport("Test - Download - FAIL")
            Form1.failTest = Form1.failTest + 1
        End If



        vis = Nothing
        project = Nothing
        Form1.logIt("Test - Download - END")
        Form1.totalTests = Form1.totalTests + 1
    End Sub

    Public Shared Sub pcmFlash()

        Form1.logIt("Test - Flash - START")
        Dim vis As Object = CreateObject("Vision.Application")
        Dim project As IATIVisionProjectInterface = vis.GetCurrentProjectInterface
        Dim pcm As IATIVisionDeviceInterface = project.FindDevice("PCM")
        Dim stat As Boolean



        Threading.Thread.Sleep(2000)
        Do

        Loop Until project.Online = True

        If project.Online = True Then
            Form1.logIt("Starting to Flash Active Strategy")
            stat = pcm.FlashActiveStrategy()
        End If

        If stat = True Then
            Form1.logIt("Test - Flash - PASS")
            Form1.logIt2("0007", "FLASH", "PASS")
            'Form1.logItReport("Test - Flash - PASS")
            Form1.passTest = Form1.passTest + 1
        Else
            Form1.logIt("Test - Flash - FAIL")
            Form1.logIt2("0007", "FLASH", "FAIL")
            'Form1.logItReport("Test - Flash - FAIL")
            Form1.failTest = Form1.failTest + 1
        End If



        vis = Nothing
        project = Nothing
        Form1.logIt("Test - Flash - END")
        Form1.totalTests = Form1.totalTests + 1
    End Sub

    Public Shared Sub pcmUpload()

        Form1.logIt("Test - Upload - START")
        Dim vis As Object = CreateObject("Vision.Application")
        Dim project As IATIVisionProjectInterface = vis.GetCurrentProjectInterface
        Dim pcm As IATIVisionDeviceInterface = project.FindDevice("PCM")
        Dim stat As Boolean
        Dim vst1, vst2 As Object


        Threading.Thread.Sleep(2000)
        Do

        Loop Until project.Online = True

        If project.Online = True Then
            Form1.logIt("Start Strategy Upload")
            stat = pcm.UploadActiveStrategy("C:\VISION Projects\Samples\VISION Demo\test_upload.vst")
        End If




        If stat = True Then
            Form1.logIt("Test - Upload - PASS")
            Form1.logIt2("0008", "UPLOAD", "PASS")
        Else
            Form1.logIt("Test - Upload - FAIL")
            Form1.logIt2("0008", "UPLOAD", "FAIL")
        End If



        vis = Nothing
        project = Nothing
        Form1.logIt("Test - Upload - END")

    End Sub
    Public Shared Function compVst()
        Dim vis As Object = CreateObject("Vision.Application")
        Dim project As IATIVisionProjectInterface = vis.GetCurrentProjectInterface
        Dim pcm As IATIVisionDeviceInterface = project.FindDevice("PCM")
        Dim same As Boolean = System.IO.File.ReadAllBytes(pcm.ActiveStrategy.FileName).SequenceEqual(System.IO.File.ReadAllBytes(pcm.ActiveStrategy.FileName))
        Return same
    End Function


    Public Shared Sub recordData()

        Form1.logIt("Test - RecordData - START")
        Dim vis As Object = CreateObject("Vision.Application")
        Dim project As IATIVisionProjectInterface = vis.GetCurrentProjectInterface
        Dim pcm As IATIVisionDeviceInterface = project.FindDevice("PCM")
        Dim screen As IATIVisionScreenInterface = project.Screens.Item(1)
        Dim rec As IATIVisionRecorderInterface = screen.FindControl("Recorder")
        Dim recfilename As String = rec.RecorderFilename
        Dim stat As Boolean

        stat = rec.Start
        If stat = False Then
            Form1.logIt("Unable to Start Recording")
            Form1.logIt("Test - RecordData - FAIL")
            Form1.failTest = Form1.failTest + 1
            Exit Sub
        Else
            Form1.logIt("Started Recording Data successfully")
            Threading.Thread.Sleep(10000)
            stat = rec.Stop
            If stat = False Then
                Form1.logIt("Unable to Stop Recording")
                Form1.logIt("Test - RecordData - FAIL")
                Form1.logIt2("0009", "RECORD DATA", "FAIL")
                Form1.failTest = Form1.failTest + 1
                Exit Sub
            Else
                Form1.logIt("Stopped Recording Data")
                Form1.logIt("Recorder file is " + rec.RecorderFileDirectoryPath + "\" + recfilename)
            End If
        End If

        If stat = True Then
            Form1.logIt("Test - RecordData - PASS")
            Form1.logIt2("0009", "RECORD DATA", "PASS")
            Form1.passTest = Form1.passTest + 1
        End If

        Form1.logIt("Test - RecordData - END")
        Form1.totalTests = Form1.totalTests + 1
    End Sub

End Class
