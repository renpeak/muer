Public Class 进度条
    Private Sub 进度条_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Enabled = True
        Timer1.Interval = 500
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If ProgressBar1.Value < ProgressBar1.Maximum Then
            ProgressBar1.Value += 20
        Else
            ProgressBar1.Value = 0
        End If
    End Sub

End Class