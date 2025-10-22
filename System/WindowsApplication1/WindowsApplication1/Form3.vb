Imports System.Data.SqlClient
Imports System.Drawing.Drawing2D

Public Class Form3
    ' Dragging
    Private isDragging As Boolean = False
    Private startPoint As Point
    Dim titleBar As New Panel()

    ' Mode flag: True = Add, False = Update
    Private isAddMode As Boolean = True

    ' Constructor for Add mode
    Public Sub New()
        InitializeComponent()
        isAddMode = True
        InitializeForm()
        SetAddMode()
    End Sub

    ' Constructor for Update mode
    Public Sub New(ByVal empName As String, ByVal empPosition As String, ByVal empRate As Decimal,
                   ByVal empAddress As String, ByVal empContact As String, ByVal empStatus As String)
        InitializeComponent()
        isAddMode = False
        InitializeForm() ' Populate ComboBoxes
        ' Assign values
        Guna2name.Text = empName
        Guna2rate.Text = empRate.ToString()
        Guna2address.Text = empAddress
        Guna2cons.Text = empContact

        ' Set Position
        If Guna2ComboBox1.Items.Contains(empPosition) Then
            Guna2ComboBox1.SelectedItem = empPosition
        Else
            Guna2ComboBox1.Items.Add(empPosition)
            Guna2ComboBox1.SelectedItem = empPosition
        End If

        ' Set Status
        If Guna2status.Items.Contains(empStatus) Then
            Guna2status.SelectedItem = empStatus
        Else
            Guna2status.Items.Add(empStatus)
            Guna2status.SelectedItem = empStatus
        End If

        ' Change button to UPDATE
        Guna2GradientButton2.Text = "UPDATE"
    End Sub

    ' Populate ComboBoxes and initialize form elements
    Private Sub InitializeForm()
        ' Populate Position ComboBox
        Guna2ComboBox1.Items.Clear()
        Guna2ComboBox1.Items.AddRange(New String() {
            "Store Manager",
            "Cashier",
            "Hardware Specialist",
            "Delivery Personnel",
            "Visual Merchandiser",
            "Stock Clerk"
        })

        ' Populate Status ComboBox
        Guna2status.Items.Clear()
        Guna2status.Items.Add("Full Time")
        Guna2status.Items.Add("Part Time")
    End Sub

    ' Set form to Add mode
    Public Sub SetAddMode()
        Guna2name.Text = ""
        Guna2ComboBox1.SelectedIndex = -1
        Guna2rate.Text = ""
        Guna2address.Text = ""
        Guna2cons.Text = ""
        Guna2status.SelectedIndex = -1
        Guna2GradientButton2.Text = "SUBMIT"
        Guna2GradientButton2.Tag = Nothing
    End Sub

    ' --- Custom title bar drag events ---
    Private Sub TitleBar_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            isDragging = True
            startPoint = e.Location
        End If
    End Sub

    Private Sub TitleBar_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
        If isDragging Then
            Dim p As Point = Me.PointToScreen(e.Location)
            Me.Location = New Point(p.X - startPoint.X, p.Y - startPoint.Y)
        End If
    End Sub

    Private Sub TitleBar_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        isDragging = False
    End Sub

    ' --- Form Load ---
    Private Sub Form3_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        ' Form settings
        Me.StartPosition = FormStartPosition.Manual
        Me.Width = 800
        Me.Height = 600
        Me.FormBorderStyle = FormBorderStyle.None
        ' Center form manually
        Me.Location = New Point(
            (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) \ 2,
            (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) \ 2
        )

        ' Apply rounded corners
        SetRoundedRegion(24)

        ' Setup title bar
        titleBar.Height = 40
        titleBar.Dock = DockStyle.Top
        titleBar.BackColor = Color.FromArgb(0, 64, 0)
        Me.Controls.Add(titleBar)

        Dim lblTitle As New Label()
        lblTitle.Text = "Employee Form"
        lblTitle.ForeColor = Color.White
        lblTitle.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblTitle.Dock = DockStyle.Fill
        lblTitle.TextAlign = ContentAlignment.MiddleCenter
        titleBar.Controls.Add(lblTitle)

        Dim btnClose As New Button()
        btnClose.Text = "X"
        btnClose.ForeColor = Color.White
        btnClose.BackColor = Color.FromArgb(0, 64, 0)
        btnClose.FlatStyle = FlatStyle.Flat
        btnClose.FlatAppearance.BorderSize = 0
        btnClose.Size = New Size(40, 40)
        btnClose.Dock = DockStyle.Right
        AddHandler btnClose.Click, Sub() Me.Close()
        titleBar.Controls.Add(btnClose)

        ' Dragging events
        AddHandler titleBar.MouseDown, AddressOf TitleBar_MouseDown
        AddHandler titleBar.MouseMove, AddressOf TitleBar_MouseMove
        AddHandler titleBar.MouseUp, AddressOf TitleBar_MouseUp
    End Sub

    ' Rounded corners
    Private Sub SetRoundedRegion(ByVal radius As Integer)
        Dim path As New Drawing2D.GraphicsPath()
        path.StartFigure()
        path.AddArc(0, 0, radius * 2, radius * 2, 180, 90)
        path.AddArc(Me.Width - radius * 2 - 1, 0, radius * 2, radius * 2, 270, 90)
        path.AddArc(Me.Width - radius * 2 - 1, Me.Height - radius * 2 - 1, radius * 2, radius * 2, 0, 90)
        path.AddArc(0, Me.Height - radius * 2 - 1, radius * 2, radius * 2, 90, 90)
        path.CloseFigure()
        Me.Region = New Region(path)
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        MyBase.OnPaint(e)
        ' Draw border
        Using pen As New Pen(Color.Black, 1)
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            Dim rect As New Rectangle(0, 0, Me.Width - 1, Me.Height - 1)
            e.Graphics.DrawArc(pen, 0, 0, 48, 48, 180, 90)
            e.Graphics.DrawArc(pen, Me.Width - 48, 0, 48, 48, 270, 90)
            e.Graphics.DrawArc(pen, Me.Width - 48, Me.Height - 48, 48, 48, 0, 90)
            e.Graphics.DrawArc(pen, 0, Me.Height - 48, 48, 48, 90, 90)
            e.Graphics.DrawLine(pen, 24, 0, Me.Width - 24, 0)
            e.Graphics.DrawLine(pen, Me.Width - 1, 24, Me.Width - 1, Me.Height - 24)
            e.Graphics.DrawLine(pen, 24, Me.Height - 1, Me.Width - 24, Me.Height - 1)
            e.Graphics.DrawLine(pen, 0, 24, 0, Me.Height - 24)
        End Using
    End Sub



    Private Sub DragForm(ByVal sender As Object, ByVal e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            Me.Location = New Point(Me.Location.X + e.X - startPoint.X, Me.Location.Y + e.Y - startPoint.Y)
        End If
    End Sub
  
    Private Sub Guna2GradientButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2GradientButton4.Click
        Me.Hide()
        ClearFields()
    End Sub
    Private Sub Guna2name_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        If Guna2name.Text.Length > 20 Then
            Guna2name.Text = Guna2name.Text.Substring(0, 20)
            Guna2name.SelectionStart = Guna2name.Text.Length  ' place cursor at end
        End If
    End Sub

    Private Sub Guna2address_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        If Guna2address.Text.Length > 20 Then
            Guna2address.Text = Guna2address.Text.Substring(0, 20)
            Guna2address.SelectionStart = Guna2address.Text.Length
        End If
    End Sub

    Private Sub Guna2cons_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        If Guna2cons.Text.Length > 20 Then
            Guna2cons.Text = Guna2cons.Text.Substring(0, 20)
            Guna2cons.SelectionStart = Guna2cons.Text.Length
        End If
    End Sub
    Private Sub Guna2GradientButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2GradientButton2.Click
        ' --- VALIDATIONS ---
        If Guna2name.Text.Trim().Length < 2 Then
            MessageBox.Show("Name must be at least 2 characters long.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Guna2address.Text.Trim().Length < 5 Then
            MessageBox.Show("Address must be at least 5 characters long.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Guna2cons.Text.Trim().Length < 7 OrElse Not IsNumeric(Guna2cons.Text) Then
            MessageBox.Show("Contact number must be at least 7 digits long and numeric.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Guna2status.Text.Trim().Length < 2 Then
            MessageBox.Show("Status must be at least 2 characters long.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If String.IsNullOrWhiteSpace(Guna2rate.Text) OrElse Not IsNumeric(Guna2rate.Text) Then
            MessageBox.Show("Please enter a valid numeric Rate Per Hour.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim connectionString As String =
        "Data Source=(LocalDB)\MSSQLLocalDB;" &
        "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
        "Integrated Security=True;"


        Using myconnection As New SqlConnection(connectionString)
            Try
                myconnection.Open()
                Dim query As String = ""

                If Guna2GradientButton2.Text = "SUBMIT" Then
                    query = "INSERT INTO Employee([Name],[Position],[RatePerHour],[Address],[Contact],[Status]) " &
                            "VALUES(@name, @position, @rate, @address, @contact, @status)"
                ElseIf Guna2GradientButton2.Text = "UPDATE" Then
                    If Guna2GradientButton2.Tag Is Nothing Then
                        MessageBox.Show("No employee selected for update.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                    query = "UPDATE Employee SET [Name]=@name, [Position]=@position, [RatePerHour]=@rate, " &
                            "[Address]=@address, [Contact]=@contact, [Status]=@status " &
                            "WHERE EmployeeID=@id"
                End If

                Using cmd As New SqlCommand(query, myconnection)
                    cmd.Parameters.AddWithValue("@name", Guna2name.Text.Trim())
                    cmd.Parameters.AddWithValue("@position", Guna2ComboBox1.Text.Trim())
                    cmd.Parameters.AddWithValue("@rate", Convert.ToDecimal(Guna2rate.Text.Trim()))
                    cmd.Parameters.AddWithValue("@address", Guna2address.Text.Trim())
                    cmd.Parameters.AddWithValue("@contact", Guna2cons.Text.Trim())
                    cmd.Parameters.AddWithValue("@status", Guna2status.Text.Trim())

                    If Guna2GradientButton2.Text = "UPDATE" Then
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(Guna2GradientButton2.Tag))
                    End If

                    cmd.ExecuteNonQuery()
                End Using

                ' --- SHOW SUCCESS MESSAGE ---
                If Guna2GradientButton2.Text = "SUBMIT" Then
                    MsgBox("Successfully added new employee", MsgBoxStyle.Information, "Success")
                Else
                    MsgBox("Successfully updated employee", MsgBoxStyle.Information, "Success")
                    Guna2GradientButton2.Text = "SUBMIT" ' Reset to Add mode
                    Guna2GradientButton2.Tag = Nothing
                End If

                ' --- REFRESH FORM2 DATAGRIDS IN REAL TIME ---
                Dim mainForm As Form2 = Application.OpenForms().OfType(Of Form2)().FirstOrDefault()
                If mainForm IsNot Nothing Then
                    mainForm.LoadEmployeeData()      ' Reload employee grid
                    mainForm.LoadAttendanceData()    ' Reload attendance if needed
                    mainForm.StyleAndAdjustDataGrid(mainForm.dgvemp)
                    mainForm.StyleAndAdjustDataGrid(mainForm.dgvtime)
                    mainForm.LoadEmployeeNames()
                    mainForm.UpdateEmployeeCount()
                    mainForm.LoadEmployeeData()   ' Refresh dgvemp immediately
                    mainForm.LoadEmployees()
                End If

                ' --- CLEAR FIELDS ---
                ClearFields()
                Me.Close()  ' Close Form3
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        End Using
        Form2.loader()
    End Sub

   

    ' --- HELPER METHOD TO CLEAR INPUT FIELDS ---
    Private Sub ClearFields()
        Guna2name.Clear()
        Guna2ComboBox1.SelectedIndex = -1
        Guna2rate.Clear()
        Guna2address.Clear()
        Guna2cons.Clear()
        Guna2status.SelectedIndex = -1
    End Sub
    Private Sub Guna2name_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Guna2name.KeyPress
        If Not Char.IsLetter(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsWhiteSpace(e.KeyChar) Then
            e.Handled = True
        Else
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub Guna2address_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Guna2address.KeyPress
        If Not Char.IsLetter(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsWhiteSpace(e.KeyChar) Then
            e.Handled = True
        Else
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub
    Private Sub Guna2cons_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Guna2cons.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub Guna2rate_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Guna2rate.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If

        ' Limit to 2 digits
        If Not Char.IsControl(e.KeyChar) AndAlso Guna2rate.Text.Length >= 2 Then
            e.Handled = True
        End If
    End Sub

End Class