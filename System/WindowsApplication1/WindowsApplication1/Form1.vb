Imports System.Net.Mail
Public Class Form1
    Private passwordVisible As Boolean = False
    Private WithEvents ToolTip1 As New ToolTip()
    Public Shared UserPassword As String = "12345"
    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_SYSCOMMAND As Integer = &H112
        Const SC_MOVE As Integer = &HF010

        ' Block moving the form
        If m.Msg = WM_SYSCOMMAND Then
            Dim command As Integer = m.WParam.ToInt32() And &HFFF0
            If command = SC_MOVE Then
                Return ' Ignore move command
            End If
        End If

        MyBase.WndProc(m)
    End Sub

    Private Sub FormLogin_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        txtpassword.PasswordChar = "*"c
        If My.Settings.IsLoggedIn Then
            Me.Hide()
            Form2.Show()
        End If

        ' Keep title bar, fixed size, no minimize/maximize
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' Center form
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) \ 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) \ 2
    End Sub
    Private Sub Guna2GradientButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2GradientButton8.Click
        Dim enteredPassword As String = txtpassword.Text.Trim()

        ' --- Check for user password ---
        If enteredPassword = UserPassword Then
            MessageBox.Show("Login successful!", "Welcome", MessageBoxButtons.OK, MessageBoxIcon.Information)

            My.Settings.IsLoggedIn = True
            My.Settings.Save()

            Me.Hide()
            Form2.Show()

            ' --- Check for admin password ---
        ElseIf enteredPassword = "admin123" Then
            MessageBox.Show("Admin access granted!", "Admin Login", MessageBoxButtons.OK, MessageBoxIcon.Information)

            My.Settings.IsLoggedIn = True
            My.Settings.Save()

            Me.Hide()
            Form2.Show()

            ' --- Invalid password ---
        Else
            MessageBox.Show("Invalid Password!", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        ' --- Clear password field ---
        txtpassword.Clear()
    End Sub


    Private Sub txtpassword_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtpassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True '
            Guna2GradientButton8.PerformClick()
        End If
    End Sub
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Guna2CustomGradientgrplogin_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Guna2CustomGradientgrplogin.Paint

    End Sub



    Private Sub txtPassword_IconRightClick(ByVal sender As Object, ByVal e As EventArgs) Handles txtpassword.IconRightClick
        passwordVisible = Not passwordVisible

        If passwordVisible Then

            txtpassword.PasswordChar = ControlChars.NullChar
            txtpassword.IconRight = My.Resources.password_hidding_icon_icon_for_data_privacy_and_sensitive_content_mark_illustration_vector_removebg_preview
        Else
            txtpassword.PasswordChar = "●"c
            txtpassword.IconRight = My.Resources.hide_svgrepo_com
        End If
    End Sub

    Private Sub txtpassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtpassword.TextChanged

    End Sub


    Private Sub Guna2HtmlLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2HtmlLabel1.Click
        MessageBox.Show("Forgot your password? Please contact the administrator to reset it.", "Forgot Password", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Guna2Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2Button1.Click

        If txtNewPassword.Text = "" Then
            MessageBox.Show("Enter a new password.")
            Exit Sub
        End If

        ' Update the shared password variable
        Form1.UserPassword = txtNewPassword.Text
        MessageBox.Show("Password has been reset successfully.", "Success")
        txtNewPassword.Clear()
        Me.Hide()
    End Sub

    Private Sub Guna2Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2Button2.Click
        AdminPanel.Visible = False
    End Sub
End Class
