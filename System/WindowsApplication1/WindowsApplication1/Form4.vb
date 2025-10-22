Imports System.Data.SqlClient
Imports System.Drawing.Drawing2D

Public Class Form4
    ' Form-level variables
    Private titleBar As Panel
    Private isDragging As Boolean = False
    Private startPoint As Point

    Public Sub New()
        InitializeComponent()
        SetupFormUI()
    End Sub

    ' ----------------------- Form UI Setup -----------------------
    Private Sub SetupFormUI()
        ' Form settings
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Width = 600
        Me.Height = 400
        Me.FormBorderStyle = FormBorderStyle.None


        ' Rounded corners
        SetRoundedRegion(24)

        ' ---------------- Title Bar ----------------
        titleBar = New Panel()
        titleBar.Height = 40
        titleBar.Dock = DockStyle.Top
        titleBar.BackColor = Color.FromArgb(0, 123, 255)
        Me.Controls.Add(titleBar)

        ' Title label
        Dim lblTitle As New Label()
        lblTitle.Text = "Edit Attendance"
        lblTitle.ForeColor = Color.White
        lblTitle.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblTitle.Dock = DockStyle.Fill
        lblTitle.TextAlign = ContentAlignment.MiddleCenter
        titleBar.Controls.Add(lblTitle)

        ' Close button
        Dim btnClose As New Button()
        btnClose.Text = "X"
        btnClose.ForeColor = Color.White
        btnClose.BackColor = Color.FromArgb(0, 123, 255)
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
        Dim path As New GraphicsPath()
        path.StartFigure()
        path.AddArc(0, 0, radius * 2, radius * 2, 180, 90)
        path.AddArc(Me.Width - radius * 2 - 1, 0, radius * 2, radius * 2, 270, 90)
        path.AddArc(Me.Width - radius * 2 - 1, Me.Height - radius * 2 - 1, radius * 2, radius * 2, 0, 90)
        path.AddArc(0, Me.Height - radius * 2 - 1, radius * 2, radius * 2, 90, 90)
        path.CloseFigure()
        Me.Region = New Region(path)
    End Sub

    ' Dragging logic
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
    Private Sub Guna2GradientButton9_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Guna2GradientButton9.Click
        If Form2.dgvtime.CurrentRow Is Nothing Then Exit Sub
        Dim row As DataGridViewRow = Form2.dgvtime.CurrentRow
        Dim recordID As Integer = CInt(row.Cells("ID").Value)

        ' Attendance Date
        Dim attendanceDate As Date = dtpWorkDate.Value.Date

        ' Convert TimeIn to 24-hour
        Dim hourIn As Integer = CInt(cmbHourIn.SelectedItem)
        Dim minuteIn As Integer = CInt(cmbMinuteIn.SelectedItem)
        Dim ampmIn As String = cmbAMPMIn.SelectedItem.ToString()
        If ampmIn = "PM" And hourIn < 12 Then hourIn += 12
        If ampmIn = "AM" And hourIn = 12 Then hourIn = 0
        Dim timeIn As New TimeSpan(hourIn, minuteIn, 0)

        ' Convert TimeOut to 24-hour
        Dim hourOut As Integer = CInt(cmbHourOut.SelectedItem)
        Dim minuteOut As Integer = CInt(cmbMinuteOut.SelectedItem)
        Dim ampmOut As String = cmbAMPMOut.SelectedItem.ToString()
        If ampmOut = "PM" And hourOut < 12 Then hourOut += 12
        If ampmOut = "AM" And hourOut = 12 Then hourOut = 0
        Dim timeOut As New TimeSpan(hourOut, minuteOut, 0)

        ' Check total working hours
        Dim totalHours As TimeSpan = timeOut - timeIn
        If totalHours.TotalHours < 0 Then totalHours += TimeSpan.FromHours(24)
        If totalHours.TotalHours > 12 Then
            MessageBox.Show("Error: Total working hours cannot exceed 12 hours!", "Invalid Time", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' --- Calculate Late Minutes ---
        Dim benchTime As New TimeSpan(8, 0, 0) ' 8:00 AM
        Dim lateMinutes As Integer = 0
        If timeIn > benchTime Then
            lateMinutes = CInt((timeIn - benchTime).TotalMinutes)
        End If

        ' --- Update database including LateMinutes ---
        Using conn As New SqlConnection(
         "Data Source=(LocalDB)\MSSQLLocalDB;" &
         "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
         "Integrated Security=True;"
     )

            Dim updateQuery As String = "UPDATE Attendance SET AttendanceDate=@date, TimeIn=@timeIn, TimeOut=@timeOut, LateMinutes=@lateMinutes WHERE ID=@id"


            Using cmd As New SqlCommand(updateQuery, conn)
                cmd.Parameters.AddWithValue("@date", attendanceDate)
                cmd.Parameters.AddWithValue("@timeIn", timeIn)
                cmd.Parameters.AddWithValue("@timeOut", timeOut)
                cmd.Parameters.AddWithValue("@lateMinutes", lateMinutes)
                cmd.Parameters.AddWithValue("@id", recordID)

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using

        ' --- Update DataGridView ---
        row.Cells("AttendanceDate").Value = attendanceDate
        row.Cells("TimeIn").Value = timeIn
        row.Cells("TimeOut").Value = timeOut
        If Form2.dgvtime.Columns.Contains("LateMinutes") Then
            row.Cells("LateMinutes").Value = lateMinutes
        End If

        MessageBox.Show("Record updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Me.Hide()
    End Sub



    Private Sub Guna2GradientButton8_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Guna2GradientButton8.Click
        Me.Hide()
    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class