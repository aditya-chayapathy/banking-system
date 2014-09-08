Public Class Welcome_Screen

    Private Sub Welcome_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ProgressBar1.Increment(5)
        If ProgressBar1.Value = ProgressBar1.Maximum Then
            Dim SAPI
            SAPI = CreateObject("SAPI.spvoice")
            SAPI.Speak("Welcome")
            Me.Hide()
            Timer1.Stop()
            Log_In_Menu.Show()
        End If

    End Sub

End Class
