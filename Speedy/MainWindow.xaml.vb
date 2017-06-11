Imports System.Net
Imports Microsoft.Win32

Public Class MainWindow
    Private WithEvents download As WebClient
    Dim done As Double
    Dim tot As Double
    Dim time As Integer

    Dim starter As Integer
    Dim stoper As Integer

    Dim start_time As Integer
    Dim stop_time As Integer

    Dim elapsed_time As TimeSpan
    Dim finpath As String




    Private Sub s_ValueChanged(ByVal sender As Object, ByVal e As RoutedPropertyChangedEventArgs(Of Double)) Handles s.ValueChanged

        Dim qw As Double
        qw = Math.Round(s.Value, 2)
        pers.Content = qw & "% Complete"
        cel.IsEnabled = True
        sp.Content = done & " MB out of " & tot & " MB"




        '---------------------------------------------------------
        If s.Value = 100 Then
            pers.Content = "Download Complete!"

            dload.IsEnabled = True
            dload.Content = "Open File"
            cel.IsEnabled = True
            cel.Content = "Finish"

        End If






    End Sub

    Public Sub finish()
        dload.Content = "Download"
        link.Text = ""
        dir.Text = ""
        cel.Content = "Cancel"
        cel.IsEnabled = False

        pers.Content = "Ready when you are!"
        sp.Content = Nothing
        s.Value = 0
        link.Text = "Enter the file link..."
        link.Foreground = Brushes.Gray

        dir.Text = "Set download location..."
        dir.Foreground = Brushes.Gray

    End Sub

    Private Sub dload_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles dload.Click
        If dload.Content = "Open File" Then

            Try
                Process.Start(finpath)
            Catch ex As Exception
                Opacity = 0.4
                MessageBox.Show("The file you're trying to open doesn't exist anymore", "File not found", MessageBoxButton.OK, MessageBoxImage.Exclamation)
                finish()
                Opacity = 1
            End Try




        Else

            Dim dlock As New Object

            SyncLock dlock

                Try
                    download.CancelAsync()
                    download.Dispose()
                Catch ex As Exception

                End Try

            End SyncLock


            Dim del As New Object

            SyncLock del


                If My.Computer.FileSystem.FileExists(dir.Text) Then
                    Opacity = 0.4
                    If MessageBox.Show("A file with the same name already exists in this location. Do you want to overwrite the file?", "Overwrite File?", MessageBoxButton.YesNo, MessageBoxImage.Question) = vbAbort.Yes Then

                        Try
                            My.Computer.FileSystem.DeleteFile(dir.Text)

                        Catch ex As Exception
                            Opacity = 0.4
                            MessageBox.Show(ex.Message)
                            Opacity = 1
                        End Try

                    Else

                        s.Value = 0
                        pers.Content = Nothing
                        sp.Content = Nothing
                        link.IsEnabled = True
                        br.IsEnabled = True
                        dir.IsEnabled = True
                        dload.IsEnabled = True
                        cel.IsEnabled = False
                        Exit Sub

                    End If
                Else

                End If
                Opacity = 1
            End SyncLock

            Try
                link.IsEnabled = False
                br.IsEnabled = False
                dir.IsEnabled = False
                dload.IsEnabled = False
                cel.IsEnabled = False


                'Download code begins here..........................

                download = New WebClient
                Dim url = link.Text
                Dim path = dir.Text
                s.Value = 0
                s.Maximum = 100

                pers.Content = "Connecting to Server..."


                If dir.Text.Contains("Desktop") Then
                    dir.Text = "Desktop"


                End If

                cel.IsEnabled = False

                finpath = path
                download.DownloadFileAsync(New Uri(url), (path))


            Catch ex1 As Exception

                Opacity = 0.4
                MessageBox.Show(ex1.Message)
                s.Value = 0
                pers.Content = Nothing
                sp.Content = Nothing
                link.IsEnabled = True
                br.IsEnabled = True
                dir.IsEnabled = True
                dload.IsEnabled = True
                cel.IsEnabled = False
                Opacity = 1
            End Try

        End If

    End Sub



    Private Sub download_DownloadProgressChanged(ByVal sender As Object, ByVal e As DownloadProgressChangedEventArgs) Handles download.DownloadProgressChanged


        done = (e.BytesReceived) / 1048576
        done = Math.Round(done, 2)
        tot = (e.TotalBytesToReceive) / 1048576
        tot = Math.Round(tot, 2)

        Dim acc As Double
        acc = (done / tot) * 100
        s.Value = acc


    End Sub





    Private Sub br_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles br.Click
        Dim dlog As New SaveFileDialog
        dlog.Title = "Select Download Location"
        dlog.ShowDialog()
        dir.Text = dlog.FileName()
        dir.Foreground = Brushes.Black


    End Sub


    Private Sub cel_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cel.Click
        If cel.Content = "Finish" Then
            finish()
        Else


            Me.Opacity = 0.4
            If MessageBox.Show("Are you sure you want to cancel the download?" & vbNewLine & vbNewLine & "NOTE: The file can become unusable", "Cancelling...", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then

                Dim path = dir.Text

                download.CancelAsync()
                s.Value = 0
                pers.Content = Nothing
                sp.Content = Nothing
                link.IsEnabled = True
                br.IsEnabled = True
                dir.IsEnabled = True
                dload.IsEnabled = True
                cel.IsEnabled = False



                Try
                    My.Computer.FileSystem.DeleteFile(path)

                Catch ex As Exception
                End Try

            Else


            End If
            Me.Opacity = 1

        End If
    End Sub

    Private Sub link_TextChanged(sender As Object, e As EventArgs) Handles link.GotFocus

        If link.Text = "Enter the file link..." Then
            link.Text = ""
            link.Foreground = Brushes.Black
        End If

    End Sub

    Private Sub lik_TextChanged(sender As Object, e As EventArgs) Handles link.LostFocus

        If link.Text = "" Then
            link.Text = "Enter the file link..."
            link.Foreground = Brushes.Gray

        End If

    End Sub


    Private Sub dir_TextChanged(sender As Object, e As EventArgs) Handles dir.GotFocus

        If dir.Text = "Set download location..." Then
            dir.Text = ""
            dir.Foreground = Brushes.Black
        End If

    End Sub

    Private Sub dire(sender As Object, e As EventArgs) Handles dir.LostFocus

        If dir.Text = "" Then
            dir.Text = "Set download location..."
            dir.Foreground = Brushes.Gray

        End If

    End Sub

End Class
