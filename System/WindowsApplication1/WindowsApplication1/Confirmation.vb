Public Class Confirmation
    Public ReadOnly Property EnteredPassword As String
        Get
            Return txtPassword.Text
        End Get
    End Property
    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_NCHITTEST As Integer = &H84
        Const HTCAPTION As Integer = 2

        MyBase.WndProc(m)
        If m.Msg = WM_NCHITTEST Then
            If m.Result = CType(HTCAPTION, IntPtr) Then
                m.Result = CType(0, IntPtr) ' Disable dragging by title bar
            End If
        End If
    End Sub
    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnOK.Click
        If String.IsNullOrWhiteSpace(txtPassword.Text) Then
            MessageBox.Show("Please enter the admin password.")
        Else
            Me.DialogResult = DialogResult.OK  ' ✅ return OK
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel  ' ❌ return Cancel
    End Sub

    Private Sub Confirmation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) \ 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) \ 2
        txtPassword.PasswordChar = "*"c
    End Sub

    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged

    End Sub
End Class
