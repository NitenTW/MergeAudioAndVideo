Imports System.ComponentModel
Imports System.IO

Public Class Form1

    Private Delegate Sub UpdataText(ByVal text As String)

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
        TextBox1.Text = GetFilePath("選擇聲音檔案", "聲音檔案|*.mp4")
    End Sub

    '選擇影像
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox2.Text = GetFilePath("選擇影像檔案", "影像檔案|*.mp4")
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
        GetTempPath()
        If IsReady() Then
            Button3.Enabled = False
            Button4.Enabled = True
            nowState = State.None
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub

    Private Sub GetTempPath()
        tmpPath.VideoPath = System.IO.Path.GetDirectoryName(TextBox2.Text) & "\"
        tmpPath.VideoFilename = System.IO.Path.GetFileName(TextBox2.Text)
        tmpPath.AudioPath = System.IO.Path.GetDirectoryName(TextBox1.Text) & "\"
        tmpPath.AudioFilename = System.IO.Path.GetFileName(TextBox1.Text)
        tmpPath.finishFileName = System.IO.Path.GetFileNameWithoutExtension(tmpPath.VideoFilename) & "_Merge.mp4"
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

        If Not My.Computer.FileSystem.FileExists(TextBox1.Text) Then
            MsgBox("聲音檔不存在")
            Return False
        End If

        If Not My.Computer.FileSystem.FileExists(TextBox2.Text) Then
            MsgBox("影像檔不存在")
            Return False
        End If

        If My.Computer.FileSystem.FileExists(tmpPath.finishFileName) Then
            MsgBox(tmpPath.finishFileName & " 已經存在，請刪除或重新命名該檔案")
            Return False
        End If

        Return True
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        nowState = State.Cancel
        If BackgroundWorker1.IsBusy Then BackgroundWorker1.CancelAsync()
        Button3.Enabled = True
        Button4.Enabled = False
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Merge()
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
            ffmpeg.Start()

            Dim output As String

            Using ffReader As StreamReader = ffmpeg.StandardError
                Do
                    If BackgroundWorker1.CancellationPending Then
                        ffmpeg.Kill()
                        Exit Do
                    End If

                    output = ffReader.ReadLine
                    Debug.WriteLine(" >>> " & output)
                    If Not output Is Nothing Then
                        Me.Invoke(New UpdataText(AddressOf UpdataUI), output)
                    End If

                Loop Until ffmpeg.HasExited And output = Nothing Or output = ""
            End Using
        End Using
    End Sub

    Private Sub UpdataUI(ByVal text As String)
        TextBox3.AppendText(text & vbNewLine)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If nowState = State.Cancel Then
            MsgBox("合併取消")
            If CheckBox1.Checked AndAlso My.Computer.FileSystem.FileExists(tmpPath.finishFileName) Then
                My.Computer.FileSystem.DeleteFile(tmpPath.finishFileName)
            End If
        Else
            MsgBox("合併完成")
            Button3.Enabled = True
            Button4.Enabled = False
            TextBox1.Text = String.Empty
            TextBox2.Text = String.Empty
            nowState = State.None
        End If
    End Sub
End Class