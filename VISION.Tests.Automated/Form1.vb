Imports System.IO
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports ServerLib
Imports System.Net.Mail
Imports Microsoft.Office.Interop

Public Class Form1
    Public Shared install_folder As String = "C:\Program Files (x86)\Accurate Technologies"
    Public Shared status As String = "Not Initialized"
    Public Shared dt As New DataTable
    Public Shared dt1 As New DataTable
    Public Shared visPath As String
    Public Shared visName As String
    Public Shared testName As String
    Public Shared totalTests As Integer = 0
    Public Shared passTest As Integer = 0
    Public Shared failTest As Integer = 0
    Public Shared reportName As String


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        dt.Columns.Add("Timestamp")
        dt.Columns.Add("Status")
        'GridView1.BestFitColumns()
        GridView4.BestFitColumns()

        dt1.Columns.Add("TEST ID")
        dt1.Columns.Add("TEST NAME")
        dt1.Columns.Add("RESULT")
        'GridView2.BestFitColumns()
        GridView3.BestFitColumns()

        NavigationFrame1.SelectedPageIndex = 0
        Dim group = New TileGroup()
        group.Text = "Vision Installs"
        TileControl1.Groups.Add(group)

        For Each folder As String In Directory.GetDirectories(install_folder)
            If File.Exists(folder + "\Vision.exe") Or File.Exists(folder + "\VISION.exe") Then
                Dim tileItem = New TileItem()
                tileItem.Elements.Add(New TileItemElement() With {.Text = Replace(folder, "C:\Program Files (x86)\Accurate Technologies\", ""), .TextAlignment = TileItemContentAlignment.TopCenter})
                group.Items.Add(tileItem)
            End If
        Next


        If Environment.GetCommandLineArgs(3).ToString = "EXECUTE" Then
            testName = Environment.GetCommandLineArgs(1).ToString
            visPath = install_folder + "\" + Replace(Environment.GetCommandLineArgs(2).ToString, "_", " ")
            NavigationFrame1.SelectedPageIndex = 3
            If testName = "QUICK_TEST" Then
                QuickTest(visPath)
                'generateReport()
                Me.Close()
            End If
        ElseIf Environment.GetCommandLineArgs(3).ToString = "arg3" Then
            'MsgBox("No commandline call")

        Else

            'NavigationFrame1.SelectedPageIndex = 0
            'Dim group = New TileGroup()
            'group.Text = "Vision Installs"
            'TileControl1.Groups.Add(group)

            'For Each folder As String In Directory.GetDirectories(install_folder)
            '    If File.Exists(folder + "\Vision.exe") Or File.Exists(folder + "\VISION.exe") Then
            '        Dim tileItem = New TileItem()
            '        tileItem.Elements.Add(New TileItemElement() With {.Text = Replace(folder, "C:\Program Files (x86)\Accurate Technologies\", ""), .TextAlignment = TileItemContentAlignment.TopCenter})
            '        group.Items.Add(tileItem)
            '    End If
            'Next
        End If
    End Sub
    Private Sub TileControl1_ItemClick(sender As Object, e As TileItemEventArgs) Handles TileControl1.ItemClick
        visPath = install_folder + "\" + e.Item.Text
        NavigationFrame1.SelectedPageIndex = 1
    End Sub

    Private Sub TileControl2_ItemClick(sender As Object, e As TileItemEventArgs) Handles TileControl2.ItemClick
        'visPath = install_folder + "\" + e.Item.Text
        If e.Item.Text = "Quick Test" Then
            testName = "QUICK_TEST"
        ElseIf e.Item.Text = "Full Test" Then
            testName = "FULL_TEST"
        ElseIf e.Item.Text = "Diagnostics" Then
            testName = "DIAGNOSTICS"
        ElseIf e.Item.Text = "Ford Dunton" Then
            testName = "FORD_DUNTON"
        ElseIf e.Item.Text = "LRW" Then
            testName = "LRW"
        ElseIf e.Item.Text = "JCB" Then
            testName = "JCB"

        End If
        NavigationFrame1.SelectedPageIndex = 2
    End Sub
    Private Sub TileControl3_ItemClick(sender As Object, e As TileItemEventArgs) Handles TileControl3.ItemClick
        If e.Item.Text = "Test Now" Then
            NavigationFrame1.SelectedPageIndex = 3
        ElseIf e.Item.Text = "Schedule" Then

        End If
    End Sub

#Region "QUICK_TEST"
    '******************************
    'Quick Test
    'Used for Quick Demo of Test Application
    'Test cases using ATI Virtual Project
    '******************************
    Public Shared Function QuickTest(vispath As String) As String
        Application.DoEvents()
        logIt("Automation Testing Started")
        visName = Replace(Replace(vispath, "C:\Program Files (x86)\Accurate Technologies\", ""), "\Vision.exe", "")
        logIt("Vision Version: " + visName)

        TestSequences.licCheck()
        TestSequences.launchVision(vispath)
        TestSequences.openProject("C:\VISION Projects\Samples\VISION Demo\Vision Demo.vpj")
        Process.Start("C:\VISION Projects\Samples\VISION Demo\XcpIpSim.exe")
        TestSequences.projectOnline()
        TestSequences.projectOffline()
        TestSequences.projectOnline()
        TestSequences.pcmDownload()
        TestSequences.pcmFlash()
        TestSequences.pcmUpload()
        TestSequences.recordData()
        TestSequences.projectOffline()

        killProcess("XcpIpSim")
        killProcess("VISION")

        generateReport()
        logIt("Automation Testing Finished")
        status = "Automation Testing Finished"
        'reportName = "C:\Vision Projects\AutomatedVisionTesting\TestReports\VisionRegressionTestReport_" + Trim(Replace(Replace(Now.ToString, "/", "_"), ":", "_")) + ".pdf"

        Return status

    End Function

    
#End Region

#Region "FULL_TEST"
    '******************************
    'Quick Test
    'Used for Quick Demo of Test Application
    'Test cases using ATI Virtual Project
    '******************************
    Public Shared Function FullTest(vispath As String) As String
        Application.DoEvents()
        'logIt("Automation Testing Started")
        'visName = Replace(Replace(vispath, "C:\Program Files (x86)\Accurate Technologies\", ""), "\Vision.exe", "")
        'logIt("Vision Version: " + visName)

        'TestSequences.licCheck()
        'TestSequences.launchVision(vispath)
        'TestSequences.openProject("C:\VISION Projects\Samples\VISION Demo\Vision Demo.vpj")
        'Process.Start("C:\VISION Projects\Samples\VISION Demo\XcpIpSim.exe")
        'TestSequences.projectOnline()
        'TestSequences.projectOffline()
        'TestSequences.projectOnline()
        'TestSequences.pcmDownload()
        'TestSequences.pcmFlash()
        'TestSequences.pcmUpload()
        'TestSequences.recordData()
        'TestSequences.projectOffline()

        'killProcess("XcpIpSim")
        'killProcess("VISION")

        'generateReport()
        'logIt("Automation Testing Finished")
        'status = "Automation Testing Finished"
        'reportName = "C:\Vision Projects\AutomatedVisionTesting\TestReports\VisionRegressionTestReport_" + Trim(Replace(Replace(Now.ToString, "/", "_"), ":", "_")) + ".pdf"

        Return status

    End Function


#End Region

#Region "DIAGNOSTICS_TEST"
    '******************************
    'Quick Test
    'Used for Quick Demo of Test Application
    'Test cases using ATI Virtual Project
    '******************************
    Public Shared Function DiagnosticsTest(vispath As String) As String
        Application.DoEvents()
        'logIt("Automation Testing Started")
        'visName = Replace(Replace(vispath, "C:\Program Files (x86)\Accurate Technologies\", ""), "\Vision.exe", "")
        'logIt("Vision Version: " + visName)

        'TestSequences.licCheck()
        'TestSequences.launchVision(vispath)
        'TestSequences.openProject("C:\VISION Projects\Samples\VISION Demo\Vision Demo.vpj")
        'Process.Start("C:\VISION Projects\Samples\VISION Demo\XcpIpSim.exe")
        'TestSequences.projectOnline()
        'TestSequences.projectOffline()
        'TestSequences.projectOnline()
        'TestSequences.pcmDownload()
        'TestSequences.pcmFlash()
        'TestSequences.pcmUpload()
        'TestSequences.recordData()
        'TestSequences.projectOffline()

        'killProcess("XcpIpSim")
        'killProcess("VISION")

        'generateReport()
        'logIt("Automation Testing Finished")
        'status = "Automation Testing Finished"
        'reportName = "C:\Vision Projects\AutomatedVisionTesting\TestReports\VisionRegressionTestReport_" + Trim(Replace(Replace(Now.ToString, "/", "_"), ":", "_")) + ".pdf"

        Return status

    End Function


#End Region


    Private Shared Sub killProcess(processname As String)
        Dim arrProcess() As Process = System.Diagnostics.Process.GetProcessesByName(processname)

        For Each p As Process In arrProcess
            p.Kill()
        Next
    End Sub

    Public Shared Sub logIt(stat As String)
        Application.DoEvents()
        Dim dr As DataRow
        dr = Form1.dt.NewRow
        dr.Item(0) = Now.ToString
        dr.Item(1) = stat
        Form1.dt.Rows.Add(dr)
        Form1.GridControl4.DataSource = Form1.dt
        Form1.GridView4.MoveLast()
        Form1.GridView4.BestFitColumns()
    End Sub

    Public Shared Sub logIt2(testID As String, testname As String, stat As String)
        Application.DoEvents()
        Dim dr1 As DataRow
        dr1 = Form1.dt1.NewRow
        dr1.Item(0) = testID
        dr1.Item(1) = testname
        dr1.Item(2) = stat
        Form1.dt1.Rows.Add(dr1)
        Form1.GridControl3.DataSource = Form1.dt1
        Form1.GridView3.MoveLast()
        Form1.GridView3.BestFitColumns()
    End Sub

    Public Shared Sub generateReport()
        Dim dir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim reportname As String = dir + "\VisionTesting\Reports\TestReport.pdf"
        'GridControl2.ExportToPdf(reportname)
        'Process.Start(reportname)
        Form1.GridControl3.ExportToPdf(reportname)
        Try
            emailReport(reportname)
        Catch ex As Exception
            logIt("Unable to email report")
        End Try
        Process.Start(reportname)
    End Sub

    Public Shared Function emailReport(reportname As String) As String

        Try

            Dim ol As New Outlook.Application()
            Dim ns As Outlook.NameSpace
            Dim fdMail As Outlook.MAPIFolder

            ns = ol.GetNamespace("MAPI")
            ns.Logon(, , True, True)

            'creating a new MailItem object
            Dim newMail As Outlook.MailItem

            'gets defaultfolder for my Outlook Outbox
            fdMail = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox)

            'assign values to the newMail MailItem
            newMail = fdMail.Items.Add(Outlook.OlItemType.olMailItem)
            newMail.Subject = "Test Report"
            newMail.Body = "This is a test e-mail message sent by an application."
            newMail.To = "sreram@ktech-international.com"
            'newMail.SaveSentMessageFolder = fdMail
            newMail.Attachments.Add(reportname)
            newMail.Send()

        Catch ex As Exception

            Throw ex

        End Try
        Return "Report Generated was Email successfully"
    End Function


    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        If NavigationFrame1.SelectedPageIndex = 1 Then
            NavigationFrame1.SelectedPageIndex = 0
        ElseIf NavigationFrame1.SelectedPageIndex = 2 Then
            NavigationFrame1.SelectedPageIndex = 1
        ElseIf NavigationFrame1.SelectedPageIndex = 3 Then
            NavigationFrame1.SelectedPageIndex = 2
        ElseIf NavigationFrame1.SelectedPageIndex = 4 Then
            NavigationFrame1.SelectedPageIndex = 3
        ElseIf NavigationFrame1.SelectedPageIndex = 5 Then
            NavigationFrame1.SelectedPageIndex = 2
        End If
    End Sub

    Private Sub NavButton3_ElementClick(sender As Object, e As DevExpress.XtraBars.Navigation.NavElementEventArgs) Handles NavButton3.ElementClick
        'Navigate to Log window
        NavigationFrame1.SelectedPageIndex = 4
    End Sub

    Private Sub TileItem8_ItemClick(sender As Object, e As TileItemEventArgs) Handles TileItem8.ItemClick
        'Navigate to Scheduler
        NavigationFrame1.SelectedPageIndex = 5
    End Sub

    Private Sub NavButton4_ElementClick(sender As Object, e As DevExpress.XtraBars.Navigation.NavElementEventArgs) Handles NavButton4.ElementClick
        If testName = "QUICK_TEST" Then
            QuickTest(visPath)
        ElseIf testName = "FULL_TEST" Then
            FullTest(visPath)
        ElseIf testName = "DIAGNOSTICS" Then
            DiagnosticsTest(visPath)
        End If
    End Sub

    Private Sub GridView3_RowCellStyle(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs) Handles GridView3.RowCellStyle
        If e.CellValue.ToString.Contains("FAIL") = True Then e.Appearance.BackColor = Color.Yellow
        If e.CellValue.ToString.Contains("PASS") = True Then e.Appearance.BackColor = Color.LightGreen
    End Sub

    Private Sub NavButton2_ElementClick(sender As Object, e As DevExpress.XtraBars.Navigation.NavElementEventArgs) Handles NavButton2.ElementClick
        generateReport()
    End Sub

    Private Sub NavButton7_ElementClick(sender As Object, e As DevExpress.XtraBars.Navigation.NavElementEventArgs) Handles NavButton7.ElementClick
        NavigationFrame1.SelectedPageIndex = 2
    End Sub
End Class
