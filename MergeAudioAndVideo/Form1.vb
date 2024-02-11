Imports System.ComponentModel

Public Class Form1

    Private Delegate Sub UpdateText(ByVal text As String)

    Private Enum State
        None = 0
        Cancel
        Finish
    End Enum

    Private Structure TempPath
        Dim VideoPath As String
        Dim AudioPath As String
        Dim VideoFilename As String
        Dim AudioFilename As String
        Dim finishFileName As String
    End Structure

    Dim nowState As State
    Dim tmpPath As TempPath

    '選擇聲音
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = GetFilePath("選擇聲音檔案", "聲音檔案|*.mp4|全部檔案|*.*")
    End Sub

    '選擇影像
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox2.Text = GetFilePath("選擇影像檔案", "影像檔案|*.mp4|全部檔案|*.*")
    End Sub

    ''' <summary>
    ''' 取得檔案路徑
    ''' </summary>
    ''' <param name="title">標題</param>
    ''' <param name="filter">檔案類型</param>
    ''' <returns>檔案路徑</returns>
    Private Function GetFilePath(ByVal title As String, ByVal filter As String) As String
        Dim result As String = String.Empty

        Using openfile As New OpenFileDialog
            openfile.Title = title
            openfile.Filter = filter
            If openfile.ShowDialog = DialogResult.OK Then
                result = openfile.FileName
            End If
        End Using

        Return result
    End Function

    '轉換
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If IsReady() Then
            GetTempPath()
            Button3.Enabled = False
            Button4.Enabled = True
            nowState = State.None

            Dim fCount As Integer = GetFrameCount()
            If Not fCount = -1 Then
                ProgressBar1.Maximum = fCount
            Else
                ProgressBar1.Maximum = 0
            End If

            GetFrameCount()
            Merge()
        End If
    End Sub

    Private Sub GetTempPath()
        tmpPath.VideoPath = System.IO.Path.GetDirectoryName(TextBox2.Text) & "\"
        tmpPath.VideoFilename = System.IO.Path.GetFileName(TextBox2.Text)
        tmpPath.AudioPath = System.IO.Path.GetDirectoryName(TextBox1.Text) & "\"
        tmpPath.AudioFilename = System.IO.Path.GetFileName(TextBox1.Text)
        tmpPath.finishFileName = GetNewFilename(System.IO.Path.GetFileNameWithoutExtension(tmpPath.VideoFilename) & "_Merge.mp4")
    End Sub

    ''' <summary>
    ''' 確認已準備好相關檔案可開始合併
    ''' </summary>
    ''' <returns>沒準備好傳回 False ，否則傳回 True</returns>
    Private Function IsReady() As Boolean
        If Not My.Computer.FileSystem.FileExists(Application.StartupPath & "\ffmpeg.exe") Then
            MsgBox("FFMPEG.exe 不存在")
            Return False
        End If

        If Not My.Computer.FileSystem.FileExists(Application.StartupPath & "\ffprobe.exe") Then
            MsgBox("FFPROBE.exe 不存在")
            Return False
        End If

        If TextBox1.Text = String.Empty Then
            MsgBox("尚未選擇聲音檔案")
            Return False
        End If

        If TextBox2.Text = String.Empty Then
            MsgBox("尚未選擇影像檔案")
            Return False
        End If

        If Not My.Computer.FileSystem.FileExists(TextBox1.Text) Then
            MsgBox("聲音檔不存在")
            Return False
        End If

        If Not My.Computer.FileSystem.FileExists(TextBox2.Text) Then
            MsgBox("影像檔不存在")
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' 取得新檔案名稱
    ''' </summary>
    ''' <param name="destinationName">要重新命名的檔案名稱 (完整路徑)</param>
    Private Function GetNewFilename(ByVal destinationName As String) As String
        Dim count As Integer = 0
        Dim tmpDirectory As String = System.IO.Path.GetDirectoryName(destinationName)
        Dim tmpMainFileName As String = System.IO.Path.GetFileNameWithoutExtension(destinationName)
        Dim tmpExtension As String = System.IO.Path.GetExtension(destinationName)
        Dim result As String = destinationName

        Do While System.IO.File.Exists(result)
            count += 1
            result = tmpDirectory & "\" & tmpMainFileName & "(" & count & ")" & tmpExtension
        Loop

        Return result
    End Function

    Private Function GetFrameCount() As Integer
        Using ffprobe As New Process
            Dim procInfo As New ProcessStartInfo

            procInfo.FileName = Application.StartupPath & "\ffprobe.exe"
            procInfo.Arguments = "-v error -select_streams v:0 -show_entries stream=nb_frames -of default=noprint_wrappers=1 " & TextBox2.Text
            procInfo.UseShellExecute = False
            procInfo.WindowStyle = ProcessWindowStyle.Hidden
            procInfo.RedirectStandardOutput = True
            procInfo.CreateNoWindow = True

            ffprobe.StartInfo = procInfo
            ffprobe.Start()

            Dim result As Integer = -1
            Dim outLine As String
            Using ffReader As System.IO.StreamReader = ffprobe.StandardOutput
                Do
                    result = GetFrameValue(ffReader.ReadLine)
                    outLine = ffReader.ReadLine
                Loop Until ffprobe.HasExited
            End Using

            ffprobe.Close()
            Return result
        End Using
    End Function

    Private Function GetFrameValue(ByVal outLine As String) As Integer
        Dim result As Integer = -1

        Dim temp As String = (Strings.Mid(outLine, 11))
        If IsNumeric(temp) Then
            result = Convert.ToInt32(temp)
        End If

        Return result
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        nowState = State.Cancel
        Button3.Enabled = True
        Button4.Enabled = False
        ProgressBar1.Value = 0
    End Sub

    ''' <summary>
    ''' 合併
    ''' </summary>
    Private Sub Merge()
        Using ffmpeg As New Process
            Dim procInfo As New ProcessStartInfo

            procInfo.FileName = Application.StartupPath & "\ffmpeg.exe"
            procInfo.Arguments = "-i """ & TextBox2.Text & """ -i """ & TextBox1.Text & """ -c:v copy -c:a aac " & tmpPath.finishFileName
            procInfo.UseShellExecute = False
            procInfo.WindowStyle = ProcessWindowStyle.Hidden
            procInfo.RedirectStandardError = True
            procInfo.CreateNoWindow = True

            ffmpeg.StartInfo = procInfo
            AddHandler ffmpeg.ErrorDataReceived, AddressOf ffmpeg_ErrorDataReceived
            AddHandler ffmpeg.Exited, AddressOf ffmpeg_Exited
            ffmpeg.Start()

            ffmpeg.BeginErrorReadLine()
            Do
                If nowState = State.Cancel Then
                    ffmpeg.Kill()
                    'Debug.WriteLine("使用者關閉")
                    Exit Do
                End If

                My.Application.DoEvents()
            Loop Until ffmpeg.HasExited


            If nowState = State.Cancel Then
                MsgBox("合併取消")
                If CheckBox1.Checked AndAlso My.Computer.FileSystem.FileExists(tmpPath.finishFileName) Then
                    My.Computer.FileSystem.DeleteFile(tmpPath.finishFileName)
                End If
            End If

            RemoveHandler ffmpeg.ErrorDataReceived, AddressOf ffmpeg_ErrorDataReceived
            RemoveHandler ffmpeg.Exited, AddressOf ffmpeg_Exited
            ffmpeg.Close()
        End Using
    End Sub

    Private Sub ffmpeg_ErrorDataReceived(sender As Object, e As DataReceivedEventArgs)
        If Not String.IsNullOrEmpty(e.Data) Then
            UpdateUI(e.Data)
        End If
    End Sub

    Private Sub UpdateUI(ByVal outLine As String)
        If Me.InvokeRequired Then
            Dim callback As New UpdateText(AddressOf UpdateUI)
            Me.Invoke(callback, outLine)
        Else
            TextBox3.AppendText(outLine & vbNewLine)

            Dim index As Integer = GetNowFrame(outLine)
            If Not ProgressBar1.Maximum = 0 AndAlso Not index = -1 Then
                ProgressBar1.Value = index
            End If
        End If
    End Sub

    Private Function GetNowFrame(ByVal outLine As String) As Integer
        Dim result As Integer = -1
        Dim index As Integer = outLine.IndexOf("fps=")

        If Not index = -1 Then
            result = Strings.Mid(outLine, 7, index - 7)
        End If

        Return result
    End Function

    Private Sub ffmpeg_Exited(sender As Object, e As EventArgs)
        MsgBox("合併完成")
        Button3.Enabled = True
        Button4.Enabled = False
        TextBox1.Text = String.Empty
        TextBox2.Text = String.Empty
        nowState = State.None
        ProgressBar1.Maximum = 0
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        nowState = State.Cancel
    End Sub
End Class