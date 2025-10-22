Module submain
    Sub Main()
        If My.Settings.IsLoggedIn Then
            Application.Run(New Form2())
        Else
            Application.Run(New Form1())
        End If
    End Sub
End Module
