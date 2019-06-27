Imports Microsoft.Win32
Imports Microsoft.Office.Interop

Public Class Form1
    Dim strTitle As String = "MultiMonitor Presenter"
    Dim BackupReg As String = ""
    Dim key As RegistryKey


    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        TextBox1.Text = OpenFileDialog1.FileName
    End Sub
    Private Sub OpenFileDialog2_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog2.FileOk
        TextBox2.Text = OpenFileDialog2.FileName
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OpenFileDialog2.ShowDialog()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        StartShow()
    End Sub

    Sub StartShow()
        On Error Resume Next

        BackupReg = GetReg()

        If BackupReg = "" Then
            MsgBox("Sorry, your system is not supported.", MsgBoxStyle.Exclamation, strTitle)
            Exit Sub
        End If

        Dim strPresentation1 As String
        Dim strPresentation2 As String
        Dim intSlideTime = 5
        Dim objPPT As New PowerPoint.Application 'Object = CreateObject("PowerPoint.Application")

        strPresentation1 = TextBox1.Text
        strPresentation2 = TextBox2.Text

        'Set to Monitor 1
        key.SetValue("DisplayMonitor", "\\.\DISPLAY1")
        'Start 1st Presentation
        objPPT.Visible = True
        Dim objPres As Object = objPPT.Presentations.Open(strPresentation1, , True)
        With objPres.SlideShowSettings
            .LoopUntilStopped = True
            .Run()
        End With

        'Set to Monitor 2
        key.SetValue("DisplayMonitor", "\\.\DISPLAY2\Monitor0")

        'Start 2nd presentation
        objPPT.Visible = True
        Dim objPres2 As PowerPoint.Presentation = objPPT.Presentations.Open(strPresentation2, , True)
        With objPres2.SlideShowSettings
            .LoopUntilStopped = True
            .Run()
        End With

        objPPT.Presentations(1).SlideShowWindow.Activate()
        objPPT.Presentations(2).SlideShowWindow.Activate()

        Err.Clear()

        Do
            System.Threading.Thread.Sleep(intSlideTime * 1000)
            objPPT.Presentations(1).SlideShowWindow.View.Next()
            objPPT.Presentations(2).SlideShowWindow.View.Next()
            If Err.Number <> 0 Then Exit Do
        Loop

        Err.Clear()
        objPres.Close()
        objPres2.Close()

        objPPT.Quit()

        'If Err.Number <> 0 Then MsgBox(Err.Number & vbCrLf & Err.Description)

        objPPT = Nothing

        objPres = Nothing
        objPres2 = Nothing

        key.SetValue("DisplayMonitor", BackupReg)


    End Sub
    Function GetReg()
        Dim tmpKey1 As String = "Software\Microsoft\Office\11.0\PowerPoint\Options"
        Dim tmpKey2 As String = "Software\Microsoft\Office\12.0\PowerPoint\Options"

        Dim val As String = ""
        Try
            key = Registry.CurrentUser.OpenSubKey(tmpKey1, True)
            val = key.GetValue("DisplayMonitor")
        Catch ex As Exception
            'Do nothing
        End Try

        If val = "" Then
            Try
                key = Registry.CurrentUser.OpenSubKey(tmpKey2, True)
                val = key.GetValue("DisplayMonitor")
            Catch ex As Exception
                'Do nothing
            End Try
        End If

        GetReg = val
    End Function
End Class
