Public Class Form2

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ProgressBar1.Value < 100 Then
            ProgressBar1.Value = ProgressBar1.Value + 5
            If ProgressBar1.Value > 20 Then
                Label1.Text = "Initializing..."
            End If
            If ProgressBar1.Value > 40 Then
                Label1.Text = "Initializing Database Connections..."
            End If
            If ProgressBar1.Value > 60 Then
                Label1.Text = "Reading Preferences..."
            End If
            If ProgressBar1.Value > 80 Then
                Label1.Text = "Starting..."
            End If
        Else
            Timer1.Stop()
            Me.Close()
        End If
    End Sub
End Class