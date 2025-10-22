Imports System.Data.SqlClient
Imports Guna.UI2.WinForms
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Threading.Tasks

Public Class Form2
    Private arrImage As Byte() = Nothing
    Dim connection As String
    Dim myconnection As SqlConnection = New SqlConnection
    Private currentEmpID As Integer
    Private currentStartDate As Date
    Private currentEndDate As Date
    Private totalLateMinutesGlobal As Integer = 0
    Dim serverTime As DateTime = DateTime.Now
    Dim lastTimeFile As String = "last_time.txt"

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' ✅ Ensure serverTime is valid
        If serverTime = Date.MinValue OrElse serverTime.Year < 1753 Then
            serverTime = DateTime.Now
        End If

        ' Advance time
        serverTime = serverTime.AddSeconds(1)

        ' Compute PH time safely
        Dim phTime As DateTime = serverTime.AddHours(8)
        timer.Text = phTime.ToString("yyyy-MM-dd HH:mm:ss")
    End Sub

    Sub clickondgv()
        Dim i As Integer
        i = dgvemp.CurrentRow.Index
    End Sub

    Private Function GetInternetTime() As DateTime
        Try
            ' --- Check internet connection first ---
            If Not My.Computer.Network.IsAvailable Then
                Throw New Exception("No internet connection")
            End If

            ' --- Create request with timeout (3 seconds) ---
            Dim request As Net.HttpWebRequest = DirectCast(Net.WebRequest.Create("http://worldclockapi.com/api/json/ph/now"), Net.HttpWebRequest)
            request.Timeout = 3000 ' 3 seconds
            request.ReadWriteTimeout = 3000

            Using response As Net.HttpWebResponse = DirectCast(request.GetResponse(), Net.HttpWebResponse)
                Using reader As New IO.StreamReader(response.GetResponseStream())
                    Dim json As String = reader.ReadToEnd()
                    Dim match As Match = Regex.Match(json, """currentDateTime"":""([^""]+)""")

                    If match.Success Then
                        Dim serverTime As DateTime = DateTime.Parse(match.Groups(1).Value)

                        ' Save last valid time
                        IO.File.WriteAllText(lastTimeFile, serverTime.ToString())
                        Return serverTime
                    End If
                End Using
            End Using

            Throw New Exception("No valid time found")
        Catch ex As Exception
            ' --- Fallback to last saved time ---
            If IO.File.Exists(lastTimeFile) Then
                Return DateTime.Parse(IO.File.ReadAllText(lastTimeFile))
            End If

            ' --- Final fallback if nothing works ---
            Return New DateTime(2000, 1, 1, 0, 0, 0)
        End Try
    End Function


    Private Sub Form2_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles Me.FormClosing
        IO.File.WriteAllText(lastTimeFile, serverTime.ToUniversalTime().ToString("o"))
        If Globals.IsLoggedIn Then
            e.Cancel = True
            Me.Hide()
        End If
    End Sub
Public Sub LoadEmployeeData()
        Using conn As New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database\payroll.mdf;Integrated Security=True;")
            Dim query As String = "SELECT EmployeeID, Name FROM Employee"


            Dim da As New SqlDataAdapter(query, conn)
            Dim dt As New DataTable()
            da.Fill(dt)

            dgvemp.DataSource = dt
            StyleAndAdjustDataGrid(dgvemp)
            dgvemployee.DataSource = dt
            StyleAndAdjustDataGrid(dgvemployee)
            dgvemployee.AllowUserToAddRows = False
            ' ✅ Hide specific columns in dgvemp (not dgvtime)
            If dgvemp.Columns.Contains("Address") Then
                dgvemp.Columns("Address").Visible = False
            End If
            If dgvemp.Columns.Contains("Contact") Then
                dgvemp.Columns("Contact").Visible = False
            End If
            If dgvemp.Columns.Contains("Status") Then
                dgvemp.Columns("Status").Visible = False
            End If
            If dgvemp.Columns.Contains("ID") Then
                dgvemp.Columns("ID").Visible = False
            End If
        End Using
    End Sub

    Public Sub LoadAttendanceData1()
        Using conn As New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database\payroll.mdf;Integrated Security=True;")
            Dim query As String = "SELECT ID, EmployeeID, EmployeeName, AttendanceDate, TimeIn, TimeOut, HalfDay, LateMinutes FROM Attendance"

            Dim da As New SqlDataAdapter(query, conn)
            Dim dt As New DataTable()
            da.Fill(dt)
            dgvtime.DataSource = dt

            StyleAndAdjustDataGrid(dgvtime, 500)
        End Using
    End Sub

    Public Sub LoadEmployeeNames()
        Using conn As New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database\payroll.mdf;Integrated Security=True;")
            ' Include LateMinutes and HalfDay in the SELECT query
            Dim query As String = "SELECT ID, EmployeeID, EmployeeName, AttendanceDate, TimeIn, TimeOut, HalfDay, LateMinutes FROM Attendance"
            Dim da As New SqlDataAdapter(query, conn)
            Dim dt As New DataTable()
            da.Fill(dt)
            dgvtime.DataSource = dt

            ' 🔒 Hide ID column
            If dgvtime.Columns.Contains("ID") Then
                dgvtime.Columns("ID").Visible = False
            End If

            ' 🔹 Optional: format LateMinutes column if needed
            If dgvtime.Columns.Contains("LateMinutes") Then
                dgvtime.Columns("LateMinutes").HeaderText = "Late Minutes"
                dgvtime.Columns("LateMinutes").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If

            ' 🔹 Optional: format HalfDay column
            If dgvtime.Columns.Contains("HalfDay") Then
                dgvtime.Columns("HalfDay").HeaderText = "Half Day"
                dgvtime.Columns("HalfDay").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If

            StyleAndAdjustDataGrid(dgvtime, 500)
            HighlightLateEmployees(dgvtime)
        End Using
    End Sub

    Private Sub LoadPayslipData()
        Dim connectionString As String =
     "Data Source=(LocalDB)\MSSQLLocalDB;" &
     "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
     "Integrated Security=True;"

        Dim query As String = "SELECT e.Name AS EmployeeName, p.GrossPay, p.NetPay, p.DateGenerated, p.PeriodStart, p.PeriodEnd FROM [Payslip] p INNER JOIN [Employee] e ON p.EmployeeID = e.EmployeeID ORDER BY p.DateGenerated DESC"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                Try
                    conn.Open()
                    Dim dt As New DataTable()
                    Dim adapter As New SqlDataAdapter(cmd)
                    adapter.Fill(dt)

                    ' ✅ Bind to your DataGridView
                    c.DataSource = dt

                    ' ✅ Apply your style
                    StyleAndAdjustDataGrid(c, 600)
                    c.ReadOnly = True

                    ' Optional: prevent column resizing and reordering
                    c.AllowUserToResizeColumns = False
                    c.AllowUserToOrderColumns = False
                Catch ex As Exception
                    MessageBox.Show("Error loading payslip data: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub
    Private Sub txtBonus_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles txtBonus.KeyPress
        ' Allow only digits, control keys (Backspace), and one decimal point
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "."c Then
            e.Handled = True
        End If

        ' Allow only one decimal point
        If e.KeyChar = "."c AndAlso txtBonus.Text.Contains(".") Then
            e.Handled = True
        End If
    End Sub

    Private Sub dgvtime_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles dgvtime.DataBindingComplete
        HighlightLateEmployees(dgvtime)
        HighlightLateEmployees(todayondutydatagrid)
    End Sub
    Public Sub loader()
        viewdata()
        LoadEmployeeNames()
        SetupCmbMonthBonus()
        SetupDgvBonus()
        LoadEmployeeBonuses()
        LoadCurrentOnDuty1()
        FixMissingTimeouts()
        LoadAttendanceData()
        LoadEmployeeData()
        LoadAttendanceData1()
        LoadEmployeeNames()
        UpdateEmployeeCount()
        UpdateOnDutyCountFromDB()
    End Sub
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblDeductionSummary1.Text =
                    "Deductions:" & Environment.NewLine &
                    "SSS = 4.5%" & "PhilHealth = 2.75%" & "Pag-Ibig = 2%"
        datetimepickerbunossorter.Format = DateTimePickerFormat.Custom
        datetimepickerbunossorter.CustomFormat = "MMMM yyyy"   ' Show month name and year
        datetimepickerbunossorter.ShowUpDown = True            ' Use up/down arrows instead of calendar

        datetimepickerbunossorter.Value = DateTime.Now
        LoadEmployeeBonuses(datetimepickerbunossorter.Value)
        SetupCmbMonthBonus()
        SetupDgvBonus()
        LoadEmployeeBonuses()
        LoadCurrentOnDuty1()
        FixMissingTimeouts()
        LoadAttendanceData()
        LoadEmployeeData()
        LoadAttendanceData1()
        txtYear.Text = DateTime.Now.Year.ToString()
        StyleAndAdjustDataGrid(dgvemp)
        StyleAndAdjustDataGrid(dgvtime)
        LoadPayslipData()
        Guna2TextBox1.MaxLength = 20
        Me.FormBorderStyle = FormBorderStyle.None
        serverTime = GetInternetTime()

        ' Save the last valid time for recovery
        IO.File.WriteAllText(lastTimeFile, serverTime.ToString("yyyy-MM-dd HH:mm:ss"))

        Timer1.Start()
        Dim lastServerTime As DateTime = DateTime.MinValue

        ' Load saved UTC time
        If IO.File.Exists(lastTimeFile) Then
            DateTime.TryParse(IO.File.ReadAllText(lastTimeFile), lastServerTime)
        End If

        ' Fetch internet time if first run
        Dim fetchedTime As DateTime = GetInternetTime()

        ' Calculate elapsed time
        If lastServerTime <> DateTime.MinValue Then
            Dim elapsed As TimeSpan = DateTime.UtcNow - lastServerTime
            serverTime = lastServerTime.Add(elapsed) ' still in UTC
        Else
            serverTime = fetchedTime.ToUniversalTime() ' convert API time to UTC
        End If

        ' Display in PH time
        timer.Text = serverTime.AddHours(8).ToString("yyyy-MM-dd HH:mm:ss")

        Timer1.Interval = 1000
        Timer1.Start()
        viewdata()
        LoadEmployeeNames()
        UpdateEmployeeCount()
        UpdateOnDutyCountFromDB()
        loader()
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) \ 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) \ 2


        For h As Integer = 0 To 23
            For m As Integer = 0 To 45 Step 15
                Dim timeStr As String = String.Format("{0:00}:{1:00}", h, m)

            Next
        Next


        dgvemp.ReadOnly = True
        dgvemp.AllowUserToAddRows = False
        Using conn As New SqlConnection(
    "Data Source=(LocalDB)\MSSQLLocalDB;" &
    "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
    "Integrated Security=True;"
)

            conn.Open()
            Dim da As New SqlDataAdapter("SELECT EmployeeID, Name FROM Employee", conn)
            Dim dt As New DataTable()
            da.Fill(dt)

            cmbEmployee.DataSource = dt
            cmbEmployee.DisplayMember = "Name"
            cmbEmployee.ValueMember = "EmployeeID"
            cmbEmployee.SelectedIndex = -1

        End Using
        dgvtime.AllowUserToAddRows = False
        dgvtime.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvtime.MultiSelect = False
        dgvtime.ReadOnly = True
        dgvtime.AutoGenerateColumns = True
        LoadAttendanceData()
        dgvtime.ClearSelection()

        cmbMonth.Items.AddRange({"January", "February", "March", "April", "May", "June",
                            "July", "August", "September", "October", "November", "December"})
        cmbMonth.SelectedIndex = Date.Now.Month - 1

        For yr As Integer = 2020 To Date.Now.Year
            cmbYear.Items.Add(yr.ToString())
        Next
        cmbYear.SelectedItem = Date.Now.Year.ToString()

        For i As Integer = 1 To 12
            Form4.cmbHourIn.Items.Add(i.ToString("D2"))
            Form4.cmbHourOut.Items.Add(i.ToString("D2"))
        Next

        ' Populate Minutes 00–59
        For i As Integer = 0 To 59
            Form4.cmbMinuteIn.Items.Add(i.ToString("D2"))
            Form4.cmbMinuteOut.Items.Add(i.ToString("D2"))
        Next

        ' Populate AM/PM
        Form4.cmbAMPMIn.Items.Add("AM")
        Form4.cmbAMPMIn.Items.Add("PM")
        Form4.cmbAMPMOut.Items.Add("AM")
        Form4.cmbAMPMOut.Items.Add("PM")
        Dim mtxt As New MaskedTextBox()
        mtxt.Mask = "00:00"
        mtxt.PromptChar = "_"   ' Show underscores for empty spaces
        Me.Controls.Add(mtxt)
    End Sub
    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_NCLBUTTONDOWN As Integer = &HA1
        Const HTCAPTION As Integer = 2

        If m.Msg = WM_NCLBUTTONDOWN AndAlso m.WParam.ToInt32() = HTCAPTION Then
            Return
        End If

        MyBase.WndProc(m)
    End Sub
    Private Sub UpdateOnDutyCountFromDB()
        Try
            Using conn As New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database\payroll.mdf;Integrated Security=True;")
                Dim query As String = "SELECT COUNT(*) FROM Attendance WHERE TimeIn IS NOT NULL AND TimeOut IS NULL"
                Using cmd As New SqlCommand(query, conn)
                    conn.Open()
                    Dim onDuty As Integer = CInt(cmd.ExecuteScalar())
                    Guna2onDuty.Text = onDuty
                End Using
            End Using
        Catch ex As Exception
            Guna2onDuty.Text = "On Duty: Error"
        End Try
    End Sub

    Public Sub UpdateEmployeeCount()
        Dim totalEmployees As Integer = dgvemp.Rows.Count
        If dgvemp.AllowUserToAddRows Then
            totalEmployees -= 1
        End If

        totalemp.Text = totalEmployees
    End Sub

    Private Sub Guna2GradientButton2_Click(ByVal sender As Object, ByVal e As EventArgs)


    End Sub


    Private Sub dgvemp_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = dgvemp.Rows(e.RowIndex)

            Dim empID As Integer = CInt(row.Cells("EmployeeID").Value)

            Form3.Guna2name.Text = row.Cells("Name").Value.ToString()
            Form3.Guna2ComboBox1.Text = row.Cells("Position").Value.ToString()
            Form3.Guna2rate.Text = row.Cells("RatePerHour").Value.ToString()
            Form3.Guna2address.Text = row.Cells("Address").Value.ToString()
            Form3.Guna2cons.Text = row.Cells("Contact").Value.ToString()
            Form3.Guna2status.Text = row.Cells("Status").Value.ToString()

            Form3.Guna2GradientButton2.Text = "UPDATE"
            Form3.Guna2GradientButton2.Tag = empID  ' store ID for update


        End If
    End Sub

    ' === CLEAR TEXTBOXES AFTER ADD/UPDATE ===
    Public Sub ClearFields1()
        Form3.Guna2name.Clear()
        Form3.Guna2ComboBox1.SelectedIndex = -1
        Form3.Guna2rate.Clear()
        Form3.Guna2address.Clear()
        Form3.Guna2cons.Clear()
        Form3.Guna2status.SelectedIndex = -1
    End Sub



    Private Sub Guna2GradientButton3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Guna2GradientButton3.Click
        ' Check if Form3 is already open
        Dim frm3 As Form3 = Application.OpenForms().OfType(Of Form3)().FirstOrDefault()

        If frm3 Is Nothing Then
            ' Not open, create new
            frm3 = New Form3()
        End If

        ' Clear fields and reset for Add mode

        frm3.Guna2GradientButton2.Text = "SUBMIT"
        frm3.Guna2GradientButton2.Tag = Nothing

        ' Show the form
        frm3.Show()
        frm3.BringToFront() ' Make sure it's visible on top
    End Sub



    Private Sub Guna2GradientButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Form3.Hide()

    End Sub




    Private Sub Guna2Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Guna2Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub Guna2GradientButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Sub viewdata()
        Dim con1 As New SqlConnection(
        "Data Source=(LocalDB)\MSSQLLocalDB;" &
        "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
        "Integrated Security=True;"
    )

        Dim sql As String = "SELECT EmployeeID, Name, Position, RatePerHour, Address, Contact, Status FROM Employee"
        Dim adapter As New SqlDataAdapter(sql, con1)
        Dim data As New DataTable
        adapter.Fill(data)
        dgvemp.DataSource = data
        Dim cmd As New SqlCommand(sql, con1)
        con1.Open()
        Dim myreader As SqlDataReader = cmd.ExecuteReader
        myreader.Read()
        con1.Close()

    End Sub
    Sub viewdata1()


    End Sub
    Private Sub Guna2GradientButton6_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Guna2GradientButton6.Click
        If cmbEmployee.SelectedIndex = -1 OrElse String.IsNullOrWhiteSpace(cmbEmployee.Text) Then
            MsgBox("Please select an employee first before Time In.", MsgBoxStyle.Exclamation, "Missing Selection")
            Exit Sub
        End If

        Dim empID As Integer
        If cmbEmployee.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbEmployee.SelectedValue.ToString(), empID) Then
            MsgBox("Invalid employee selection. Please refresh the list.", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If

        Dim empName As String = cmbEmployee.Text

        ' --- Get current time ---
        Dim phTime As DateTime = serverTime.AddHours(8) ' adjust if needed
        Dim todayDate As Date = phTime.Date

        ' --- Define bench in and last allowed time in ---
        Dim benchIn As DateTime = todayDate.AddHours(7) ' 7:00 AM start
        Dim lastAllowedTimeIn As DateTime = todayDate.AddHours(14) ' cannot time in after 2 PM

        If phTime >= lastAllowedTimeIn Then
            MsgBox("Time-in is no longer allowed after 2:00 PM.", MsgBoxStyle.Exclamation, "Too Late")
            Exit Sub
        End If

        ' --- Determine actual TimeIn ---
        Dim actualTimeIn As DateTime = If(phTime < benchIn, benchIn, phTime)

        ' --- Check for half-day ---
        Dim halfDay As Boolean = (phTime.Hour = 12 OrElse phTime.Hour = 13)

        ' --- Calculate late minutes ---
        Dim lateMinutes As Integer = 0
        If actualTimeIn > benchIn Then
            lateMinutes = CInt((actualTimeIn - benchIn).TotalMinutes)
        End If

        Using conn As New SqlConnection(
            "Data Source=(LocalDB)\MSSQLLocalDB;" &
            "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
            "Integrated Security=True;"
        )

            conn.Open()

            ' --- Check if already timed in ---
            Dim checkQuery As String = "SELECT COUNT(*) FROM Attendance WHERE EmployeeID=@empID AND AttendanceDate=@date"
            Using checkCmd As New SqlCommand(checkQuery, conn)
                checkCmd.Parameters.AddWithValue("@empID", empID)
                checkCmd.Parameters.AddWithValue("@date", todayDate)
                Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                If count > 0 Then
                    MsgBox("This employee already timed in today!", MsgBoxStyle.Exclamation, "Already Timed In")
                    Exit Sub
                End If
            End Using

            ' --- Insert attendance record with late minutes ---
            Dim insertQuery As String = "INSERT INTO Attendance (EmployeeID, EmployeeName, AttendanceDate, TimeIn, HalfDay, LateMinutes) " &
                                        "VALUES (@empID, @empName, @date, @timeIn, @halfDay, @lateMinutes)"
            Using insertCmd As New SqlCommand(insertQuery, conn)
                insertCmd.Parameters.AddWithValue("@empID", empID)
                insertCmd.Parameters.AddWithValue("@empName", empName)
                insertCmd.Parameters.AddWithValue("@date", todayDate)
                insertCmd.Parameters.AddWithValue("@timeIn", actualTimeIn.TimeOfDay)
                insertCmd.Parameters.AddWithValue("@halfDay", halfDay)
                insertCmd.Parameters.AddWithValue("@lateMinutes", lateMinutes)
                insertCmd.ExecuteNonQuery()
            End Using
        End Using

        Dim msg As String = "Time In saved for " & empName & " at " & actualTimeIn.ToString("hh:mm tt")
        If halfDay Then msg &= " (Marked as Half Day)"
        If lateMinutes > 0 Then msg &= " (Late: " & lateMinutes & " mins)"
        MsgBox(msg, MsgBoxStyle.Information, "Success")

        LoadAttendanceData1()
        UpdateOnDutyCountFromDB()
        cmbEmployee.SelectedIndex = -1
        LoadAttendanceData()
        loader()
    End Sub


    Private Sub Guna2GradientButton7_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Guna2GradientButton7.Click
        If cmbEmployee.SelectedIndex = -1 Then
            MsgBox("Please select an employee.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim empID As Integer = Convert.ToInt32(cmbEmployee.SelectedValue)
        Dim empName As String = cmbEmployee.Text

        ' Convert server time (UTC) to PH time
        Dim phTime As DateTime = serverTime.AddHours(8)
        Dim todayDate As Date = phTime.Date

        ' ✅ Define official bench out time (6:00 PM)
        Dim benchOut As DateTime = todayDate.AddHours(18)

        ' ✅ Determine final TimeOut to record
        Dim actualTimeOut As DateTime
        If phTime < benchOut Then
            ' Before 6 PM → normal Time Out
            actualTimeOut = phTime
        Else
            ' 6 PM or later → still record actual time for overtime calc
            actualTimeOut = phTime
        End If

        ' Prepare to calculate overtime (in minutes)
        Dim overtimeMinutes As Integer = 0
        If actualTimeOut > benchOut Then
            overtimeMinutes = CInt((actualTimeOut - benchOut).TotalMinutes)
        End If
        Using conn As New SqlConnection(
            "Data Source=(LocalDB)\MSSQLLocalDB;" &
            "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
            "Integrated Security=True;"
        )

            conn.Open()

            ' Check if Time In exists and Time Out is NULL
            Dim checkQuery As String = "SELECT COUNT(*) FROM Attendance " &
                                       "WHERE EmployeeID=@empID " &
                                       "AND AttendanceDate=@date " &
                                       "AND TimeIn IS NOT NULL " &
                                       "AND TimeOut IS NULL"

            Using checkCmd As New SqlCommand(checkQuery, conn)
                checkCmd.Parameters.AddWithValue("@empID", empID)
                checkCmd.Parameters.AddWithValue("@date", todayDate)

                Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                If count = 0 Then
                    MsgBox("No Time In record found for today. Cannot Time Out.", MsgBoxStyle.Exclamation)
                    Exit Sub
                End If
            End Using

            ' ✅ Update Time Out + Overtime if applicable
            Dim query As String = "UPDATE Attendance SET TimeOut=@timeOut, OvertimeMinutes=@ot " &
                                  "WHERE EmployeeID=@empID AND AttendanceDate=@date AND TimeOut IS NULL"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@empID", empID)
                cmd.Parameters.AddWithValue("@date", todayDate)
                cmd.Parameters.AddWithValue("@timeOut", actualTimeOut.TimeOfDay)
                cmd.Parameters.AddWithValue("@ot", overtimeMinutes)
                cmd.ExecuteNonQuery()
            End Using
        End Using

        ' ✅ Message feedback
        If overtimeMinutes > 0 Then
            MsgBox("Time Out saved for " & empName & vbCrLf &
                   "Overtime: " & overtimeMinutes & " minute(s)", MsgBoxStyle.Information, "Success")
        Else
            MsgBox("Time Out saved for " & empName & " at " & actualTimeOut.ToString("hh:mm tt"), MsgBoxStyle.Information, "Success")
        End If
        LoadAttendanceData1()
        UpdateOnDutyCountFromDB()
        cmbEmployee.SelectedIndex = -1
        LoadAttendanceData()
        loader()
    End Sub


    Public Sub LoadEmployees()
        Using conn As New SqlConnection(
         "Data Source=(LocalDB)\MSSQLLocalDB;" &
         "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
         "Integrated Security=True;"
     )

            conn.Open()

            Dim query As String = "SELECT EmployeeID, Name FROM Employee"
            Using da As New SqlDataAdapter(query, conn)
                Dim dt As New DataTable()
                da.Fill(dt)

                cmbEmployee.DataSource = dt
                cmbEmployee.DisplayMember = "Name"
                cmbEmployee.ValueMember = "EmployeeID"
                cmbEmployee.SelectedIndex = -1
            End Using
        End Using
    End Sub


    Public Sub LoadAttendanceData()
        Dim connectionString As String =
     "Data Source=(LocalDB)\MSSQLLocalDB;" &
     "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
     "Integrated Security=True;"

        Dim dt As New DataTable()


        Dim selectedID As Integer = -1
        If dgvtime.CurrentRow IsNot Nothing Then
            selectedID = CInt(dgvtime.CurrentRow.Cells("ID").Value)
        End If

        Using myconnection As New SqlConnection(connectionString)
            myconnection.Open()
            Dim query As String = "SELECT TOP (1000) [ID], [EmployeeID], [EmployeeName], [AttendanceDate], [TimeIn], [TimeOut] " &
                                  "FROM [Attendance] " &
                                  "ORDER BY [AttendanceDate] DESC, [TimeIn] DESC"
            Using cmd As New SqlCommand(query, myconnection)
                Using adapter As New SqlDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using

        dgvtime.DataSource = dt


        If selectedID <> -1 Then
            For Each row As DataGridViewRow In dgvtime.Rows
                If CInt(row.Cells("ID").Value) = selectedID Then
                    row.Selected = True
                    dgvtime.FirstDisplayedScrollingRowIndex = row.Index
                    Exit For
                End If
            Next
        End If
    End Sub
    Private Sub Guna2GradientButton9_Click(ByVal sender As Object, ByVal e As EventArgs)


    End Sub




    Private Sub ComputeTotalHours()
        Dim totalHours As Double = 0

        For Each row As DataGridViewRow In dgvtime.Rows
            If row.IsNewRow Then Continue For

            Dim timeIn As TimeSpan
            Dim timeOut As TimeSpan

            If row.Cells("TimeIn").Value IsNot Nothing AndAlso row.Cells("TimeOut").Value IsNot Nothing Then
                timeIn = CType(row.Cells("TimeIn").Value, TimeSpan)
                timeOut = CType(row.Cells("TimeOut").Value, TimeSpan)

                Dim worked As Double = (timeOut - timeIn).TotalHours

                If worked < 0 Then
                    worked += 24
                End If

                totalHours += worked
            End If
        Next

    End Sub


    Private Function GetTimeValue(ByVal val As Object) As TimeSpan
        If TypeOf val Is TimeSpan Then
            Return DirectCast(val, TimeSpan)
        ElseIf TypeOf val Is DateTime Then
            Return DirectCast(val, DateTime).TimeOfDay
        ElseIf val IsNot Nothing Then
            Dim ts As TimeSpan
            If TimeSpan.TryParse(val.ToString(), ts) Then
                Return ts
            End If
        End If
        Return TimeSpan.Zero
    End Function
    Private Sub Guna2GradientButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)



    End Sub







    Private Sub Guna2GradientButton13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GroupBoxPayslip.Visible = False
    End Sub
    Private Sub Guna2GradientButton15_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Guna2GradientButton15.Click
        ' First, check something is selected
        If dgvtime.CurrentRow Is Nothing Then
            MsgBox("Please select a record to edit.", MsgBoxStyle.Exclamation, "No Selection")
            Exit Sub
        End If

        ' Show confirmation dialog
        Using confirm As New Confirmation()
            If confirm.ShowDialog(Me) = DialogResult.OK Then
                ' Replace "12345" with your actual admin password
                If confirm.EnteredPassword = "12345" Then

                    ' If the password is correct, proceed with loading into the edit panel
                    Dim row As DataGridViewRow = dgvtime.CurrentRow

                    Form4.Show()

                    ' Load Attendance Date
                    If Not IsDBNull(row.Cells("AttendanceDate").Value) Then
                        Form4.dtpWorkDate.Value = CDate(row.Cells("AttendanceDate").Value)
                    Else
                        Form4.dtpWorkDate.Value = Date.Today
                    End If

                    ' Load TimeIn
                    If Not IsDBNull(row.Cells("TimeIn").Value) Then
                        Dim tsIn As TimeSpan = TimeSpan.Parse(row.Cells("TimeIn").Value.ToString())
                        Dim dtIn As New DateTime(2000, 1, 1, tsIn.Hours, tsIn.Minutes, 0)
                        Form4.cmbHourIn.SelectedItem = dtIn.ToString("hh")
                        Form4.cmbMinuteIn.SelectedItem = dtIn.ToString("mm")
                        Form4.cmbAMPMIn.SelectedItem = dtIn.ToString("tt")
                    End If

                    ' Load TimeOut
                    If Not IsDBNull(row.Cells("TimeOut").Value) Then
                        Dim tsOut As TimeSpan = TimeSpan.Parse(row.Cells("TimeOut").Value.ToString())
                        Dim dtOut As New DateTime(2000, 1, 1, tsOut.Hours, tsOut.Minutes, 0)
                        Form4.cmbHourOut.SelectedItem = dtOut.ToString("hh")
                        Form4.cmbMinuteOut.SelectedItem = dtOut.ToString("mm")
                        Form4.cmbAMPMOut.SelectedItem = dtOut.ToString("tt")
                    End If

                Else
                    MessageBox.Show("Invalid admin password.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End Using

    End Sub
    Private Sub dgvpayslip_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles c.SelectionChanged

    End Sub


    Private Sub FixMissingTimeouts()
        Dim currentPHTime As DateTime = serverTime.AddHours(8)
        Dim todayDate As Date = currentPHTime.Date
        Dim autoTimeout As DateTime = todayDate.AddHours(18.5) ' 6:30 PM

        ' ✅ Only run if it's already past 6:30 PM
        If currentPHTime < autoTimeout Then Exit Sub

        Using conn As New SqlConnection(
     "Data Source=(LocalDB)\MSSQLLocalDB;" &
     "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
     "Integrated Security=True;"
 )

            conn.Open()

            ' ✅ Find all employees who timed in but didn’t time out today
            Dim updateQuery As String =
                "UPDATE Attendance " &
                "SET TimeOut = @autoOut, OvertimeMinutes = 30 " &
                "WHERE AttendanceDate = @date " &
                "AND TimeOut IS NULL " &
                "AND TimeIn IS NOT NULL"

            Using cmd As New SqlCommand(updateQuery, conn)
                cmd.Parameters.AddWithValue("@autoOut", autoTimeout.TimeOfDay)
                cmd.Parameters.AddWithValue("@date", todayDate)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub



    Private Sub Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub dgvtime_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub cmbEmployee_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub Guna2GradientButton12_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Guna2GradientButton12.Click
        If dgvtime.SelectedRows.Count = 0 Then
            MsgBox("Please select a record to delete.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim selectedRow As DataGridViewRow = dgvtime.SelectedRows(0)
        Dim attendanceID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value)
        Dim empName As String = selectedRow.Cells("EmployeeName").Value.ToString()


        Dim confirm As New Confirmation()
        If confirm.ShowDialog(Me) = DialogResult.OK Then
            If confirm.EnteredPassword = "12345" Then

                Dim connectionString As String =
     "Data Source=(LocalDB)\MSSQLLocalDB;" &
     "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
     "Integrated Security=True;"

                Using conn As New SqlConnection(connectionString)
                    conn.Open()
                    Using cmd As New SqlCommand("DELETE FROM Attendance WHERE ID=@id", conn)
                        cmd.Parameters.AddWithValue("@id", attendanceID)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using

                MsgBox("Attendance record deleted successfully for " & empName, MsgBoxStyle.Information)

                Dim rowIndex As Integer = selectedRow.Index
                LoadAttendanceData()

                If dgvtime.Rows.Count > 0 Then
                    rowIndex = Math.Min(rowIndex, dgvtime.Rows.Count - 1)
                    dgvtime.Rows(rowIndex).Selected = True
                    dgvtime.FirstDisplayedScrollingRowIndex = rowIndex
                End If
            Else
                MessageBox.Show("Invalid admin password.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub




    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub Guna2Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Guna2GradientButton10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2GradientButton10.Click
        GroupBoxPayslip.Visible = False

    End Sub

    Private Sub dgvemp_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub
    Private Sub Guna2GradientButton13_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2GradientButton13.Click

    End Sub
    Private Sub Preview_Shown(ByVal sender As Object, ByVal e As EventArgs)
        Dim ppd As PrintPreviewDialog = CType(sender, PrintPreviewDialog)

        ' Print silently
        ppd.Document.PrintController = New Printing.StandardPrintController()
        ppd.Document.Print()

        ' Close the preview
        ppd.Close()

        ' Show success message and hide panel
        MessageBox.Show("Successfully exported PDF file!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        GroupBoxPayslip.Visible = False
    End Sub
    Private Sub SavePayslip(ByVal empID As Integer, ByVal startDate As Date, ByVal endDate As Date,
                            ByVal totalDays As Integer, ByVal totalHours As Decimal,
                            ByVal rate As Decimal, ByVal gross As Decimal, ByVal sss As Decimal,
                            ByVal philhealth As Decimal, ByVal pagibig As Decimal)

        ' ✅ Auto-calculate NetPay
        Dim net As Decimal = gross - (sss + philhealth + pagibig)

        Using conn As New SqlConnection(
     "Data Source=(LocalDB)\MSSQLLocalDB;" &
     "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
     "Integrated Security=True;"
 )

            conn.Open()

            ' --- Check if Employee exists
            Dim checkEmployeeQuery As String = "SELECT COUNT(*) FROM Employee WHERE EmployeeID=@empID"
            Using checkEmpCmd As New SqlCommand(checkEmployeeQuery, conn)
                checkEmpCmd.Parameters.AddWithValue("@empID", empID)
                Dim count As Integer = CInt(checkEmpCmd.ExecuteScalar())
                If count = 0 Then
                    MessageBox.Show("Employee ID does not exist!")
                    Exit Sub
                End If
            End Using

            ' --- Check if Payslip already exists for this period
            Dim checkPayslipQuery As String = "SELECT COUNT(*) FROM Payslip " &
                                             "WHERE EmployeeID=@empID AND PeriodStart=@start AND PeriodEnd=@end"
            Using checkPayslipCmd As New SqlCommand(checkPayslipQuery, conn)
                checkPayslipCmd.Parameters.AddWithValue("@empID", empID)
                checkPayslipCmd.Parameters.AddWithValue("@start", startDate)
                checkPayslipCmd.Parameters.AddWithValue("@end", endDate)
                Dim existing As Integer = CInt(checkPayslipCmd.ExecuteScalar())
                If existing > 0 Then
                    MessageBox.Show("Payslip already exists for this period!")
                    Exit Sub
                End If
            End Using

            ' --- Insert Payslip
            Dim insertQuery As String = "INSERT INTO Payslip " &
                "(EmployeeID, PeriodStart, PeriodEnd, TotalDaysWorked, TotalHoursWorked, RatePerHour, " &
                "GrossPay, SSS, PhilHealth, PagIbig, NetPay) " &
                "VALUES (@empID, @start, @end, @days, @hours, @rate, @gross, @sss, @philhealth, @pagibig, @net)"

            Using insertCmd As New SqlCommand(insertQuery, conn)
                insertCmd.Parameters.AddWithValue("@empID", empID)
                insertCmd.Parameters.AddWithValue("@start", startDate)
                insertCmd.Parameters.AddWithValue("@end", endDate)
                insertCmd.Parameters.AddWithValue("@days", totalDays)
                insertCmd.Parameters.AddWithValue("@hours", totalHours)
                insertCmd.Parameters.AddWithValue("@rate", rate)
                insertCmd.Parameters.AddWithValue("@gross", gross)
                insertCmd.Parameters.AddWithValue("@sss", sss)
                insertCmd.Parameters.AddWithValue("@philhealth", philhealth)
                insertCmd.Parameters.AddWithValue("@pagibig", pagibig)
                insertCmd.Parameters.AddWithValue("@net", net)

                insertCmd.ExecuteNonQuery()
            End Using
        End Using

        MessageBox.Show("Payslip saved successfully!")
    End Sub

    Private Sub HighlightLateEmployees(ByVal dgv As DataGridView)
        Try
            Dim officialStart As TimeSpan = New TimeSpan(8, 0, 0) ' 8:00 AM

            For Each row As DataGridViewRow In dgv.Rows
                If row.IsNewRow Then Continue For

                Dim cellValue = row.Cells("TimeIn").Value
                If cellValue Is Nothing OrElse IsDBNull(cellValue) Then Continue For

                Dim timeIn As TimeSpan

                If TypeOf cellValue Is TimeSpan Then
                    timeIn = DirectCast(cellValue, TimeSpan)
                ElseIf TypeOf cellValue Is String Then
                    Dim parsed As TimeSpan
                    If TimeSpan.TryParse(cellValue, parsed) Then
                        timeIn = parsed
                    Else
                        Continue For
                    End If
                ElseIf TypeOf cellValue Is DateTime Then
                    timeIn = DirectCast(cellValue, DateTime).TimeOfDay
                Else
                    Continue For
                End If

                If timeIn > officialStart Then
                    row.DefaultCellStyle.BackColor = Color.LightCoral
                Else
                    row.DefaultCellStyle.BackColor = Color.LightGreen
                End If
            Next
        Catch ex As Exception
            MsgBox("Error highlighting: " & ex.Message)
        End Try
    End Sub

    Public Sub StyleAndAdjustDataGrid(ByVal dgv As DataGridView, Optional ByVal respectDesignerSize As Boolean = True)
        ' --- Ensure headers are visible ---
        dgv.ColumnHeadersVisible = True
        dgv.EnableHeadersVisualStyles = False

        ' --- General style ---
        With dgv
            .BorderStyle = BorderStyle.FixedSingle
            .BackgroundColor = Color.White
            .GridColor = Color.Black
            .RowHeadersVisible = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .AllowUserToResizeRows = False
            .AllowUserToResizeColumns = False
        End With

        ' --- Column header style ---
        With dgv.ColumnHeadersDefaultCellStyle
            .BackColor = Color.Navy
            .ForeColor = Color.White
            .Font = New Font("Arial", 10, FontStyle.Bold)
            .Alignment = DataGridViewContentAlignment.MiddleCenter
        End With

        ' --- Row style ---
        With dgv.RowsDefaultCellStyle
            .BackColor = Color.White
            .ForeColor = Color.Black
            .SelectionBackColor = Color.DodgerBlue
            .SelectionForeColor = Color.White
        End With

        ' --- Alternating rows style ---
        With dgv.AlternatingRowsDefaultCellStyle
            .BackColor = Color.LightGray
            .ForeColor = Color.Black
        End With

        ' --- Adjust height only if not respecting designer size ---
        If respectDesignerSize = False Then
            Dim totalHeight As Integer = dgv.ColumnHeadersHeight
            For Each row As DataGridViewRow In dgv.Rows
                If Not row.IsNewRow Then
                    totalHeight += row.Height
                End If
            Next
            dgv.Height = totalHeight + 2
        End If
    End Sub
    Private Sub Guna2GradientButton14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2GradientButton14.Click
        Panel1.Visible = True
        Panel2.Visible = False
        cmbEmployee.Text = ""
        Panel6.Visible = False
        bunospanel.Visible = False
    End Sub

    Private Sub Guna2GradientButton16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2GradientButton16.Click
        Panel1.Visible = False
        Panel2.Visible = True
        cmbEmployee.Text = ""
        Panel6.Visible = False
        bunospanel.Visible = False
    End Sub

    Private Sub Guna2GradientButton17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2GradientButton17.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to logout?", "Confirm Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            ' Clear login settings
            My.Settings.IsLoggedIn = False
            My.Settings.Save()

            ' Close all open forms except the login form
            For Each f As Form In Application.OpenForms.Cast(Of Form)().ToList()
                If f.Name <> "Form1" Then
                    f.Close()
                End If
            Next

            ' Show login form
            Form1.Show()
        End If
    End Sub



    Private Sub Guna2TextBox1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2TextBox1.TextChanged
        Dim searchText As String = Guna2TextBox1.Text.Trim()

        Using conn As New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database\payroll.mdf;Integrated Security=True;")
            Dim query As String

            If String.IsNullOrEmpty(searchText) Then
                query = "SELECT EmployeeID, Name, Position, RatePerHour, Address, Contact, Status FROM Employee"
            Else
                query = "SELECT EmployeeID, Name, Position, RatePerHour, Address, Contact, Status " &
                        "FROM Employee " &
                        "WHERE Name LIKE @search OR Position LIKE @search OR CAST(EmployeeID AS NVARCHAR) LIKE @search"
            End If

            Using cmd As New SqlCommand(query, conn)
                If Not String.IsNullOrEmpty(searchText) Then
                    cmd.Parameters.AddWithValue("@search", "%" & searchText & "%")
                End If

                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                dgvemp.DataSource = dt
            End Using
        End Using
    End Sub
' ================================
    ' Button Click: Calculate and Print
    ' ================================
    Private Sub Guna2GradientButton1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2GradientButton1.Click
        ' --- Check if employee selected ---
        If dgvemp.CurrentRow Is Nothing Then
            MessageBox.Show("Please select an employee first.", "No Employee Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim empID As Integer = CInt(dgvemp.CurrentRow.Cells("EmployeeID").Value)
        Dim hourlyRate As Decimal = 0D
        Dim empPosition As String = ""
        Dim empName As String = ""

        ' --- Get Employee Info ---
        Using conn As New SqlConnection(
      "Data Source=(LocalDB)\MSSQLLocalDB;" &
      "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
      "Integrated Security=True;"
  )

            Dim query As String = "SELECT Name, RatePerHour, Position FROM Employee WHERE EmployeeID=@id"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", empID)
                conn.Open()
                Using rdr As SqlDataReader = cmd.ExecuteReader()
                    If rdr.Read() Then
                        If Not IsDBNull(rdr("Name")) Then empName = rdr("Name").ToString()
                        If Not IsDBNull(rdr("RatePerHour")) Then hourlyRate = Convert.ToDecimal(rdr("RatePerHour"))
                        If Not IsDBNull(rdr("Position")) Then empPosition = rdr("Position").ToString()
                    Else
                        Exit Sub
                    End If
                End Using
            End Using
        End Using

        ' --- Validate Month & Year ---
        If cmbMonth.SelectedIndex = -1 Then
            MessageBox.Show("Please select a month first.")
            Exit Sub
        End If
        If cmbYear.SelectedIndex = -1 Then
            MessageBox.Show("Please select a year first.")
            Exit Sub
        End If

        Dim year As Integer = CInt(cmbYear.SelectedItem)
        Dim month As Integer = cmbMonth.SelectedIndex + 1
        Dim startDate As New Date(year, month, 1)
        Dim endDate As New Date(year, month, Date.DaysInMonth(year, month))

        ' --- Calculate attendance ---
        Dim totalHours As Double = 0
        Dim totalDays As Integer = 0
        Dim totalLateMinutes As Double = 0

        Using conn As New SqlConnection(
     "Data Source=(LocalDB)\MSSQLLocalDB;" &
     "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
     "Integrated Security=True;"
 )

            Dim query As String = "SELECT AttendanceDate, TimeIn, TimeOut FROM Attendance " &
                                  "WHERE EmployeeID=@id AND AttendanceDate BETWEEN @start AND @end"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", empID)
                cmd.Parameters.AddWithValue("@start", startDate)
                cmd.Parameters.AddWithValue("@end", endDate)
                conn.Open()
                Using rdr As SqlDataReader = cmd.ExecuteReader()
                    Dim dutyDates As New HashSet(Of Date)()
                    While rdr.Read()
                        Dim timeIn As TimeSpan = TimeSpan.Zero
                        Dim timeOut As TimeSpan = TimeSpan.Zero

                        If Not IsDBNull(rdr("TimeIn")) Then timeIn = TimeSpan.Parse(rdr("TimeIn").ToString())
                        If Not IsDBNull(rdr("TimeOut")) Then timeOut = TimeSpan.Parse(rdr("TimeOut").ToString())

                        ' Worked hours
                        Dim worked As Double = (timeOut - timeIn).TotalHours
                        If worked < 0 Then worked += 24
                        totalHours += worked

                        ' Late calculation (shift starts at 8:00 AM)
                        Dim scheduledIn As TimeSpan = TimeSpan.Parse("08:00:00")
                        If timeIn > scheduledIn Then
                            totalLateMinutes += (timeIn - scheduledIn).TotalMinutes
                        End If

                        dutyDates.Add(CType(rdr("AttendanceDate"), Date))
                    End While
                    totalDays = dutyDates.Count
                End Using
            End Using
        End Using

        If totalDays = 0 Then
            MessageBox.Show("No attendance records found for " & empName & " in " & cmbMonth.SelectedItem & " " & year.ToString(),
                            "No Records", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        ' --- Get Bonus ---
        Dim bonus As Decimal = 0
        Using conn As New SqlConnection(
            "Data Source=(LocalDB)\MSSQLLocalDB;" &
            "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
            "Integrated Security=True;"
        )

            Dim query As String = "SELECT TOP 1 BonusAmount FROM EmployeeBonus " &
                                  "WHERE EmployeeID=@id AND BonusMonth=@month AND BonusYear=@year " &
                                  "ORDER BY DateAdded DESC"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", empID)
                cmd.Parameters.AddWithValue("@month", month)
                cmd.Parameters.AddWithValue("@year", year)
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing Then bonus = Convert.ToDecimal(result)
            End Using
        End Using

        ' --- Compute gross ---
        Dim gross As Decimal = CDec(totalHours) * hourlyRate

        ' --- Late deduction ---
        Dim lateDeduction As Decimal = (totalLateMinutes / 60D) * hourlyRate
        Dim otherDeduction As Decimal = 0

        ' --- Mandatory contributions ---
        ' --- Mandatory contributions ---
        Dim sssRate As Decimal = 0.045D
        Dim philhealthRate As Decimal = 0.0275D
        Dim pagibigRate As Decimal = 0.02D

        Dim sssMin As Decimal = 90D
        Dim philhealthMin As Decimal = 45D
        Dim pagibigMax As Decimal = 100D

        Dim sss As Decimal = Math.Max(gross * sssRate, sssMin)
        Dim philhealth As Decimal = Math.Max(gross * philhealthRate, philhealthMin)
        Dim pagibig As Decimal = Math.Min(gross * pagibigRate, pagibigMax)

        ' --- Compute withholding tax ---
        Dim tax As Decimal = ComputeWithholdingTax(gross + bonus, sss, philhealth, pagibig)

        ' --- Compute net pay ---
        Dim totalDeduction As Decimal = sss + philhealth + pagibig + lateDeduction + otherDeduction + tax
        Dim netPay As Decimal = gross + bonus - totalDeduction
        ' --- Adjust totalHours: only count time from 8:00 AM onward ---
       

        ' --- Update Labels ---
        lblEmpID.Text = "Employee ID: " & empID
        lblEmpName.Text = "Employee Name: " & empName
        lblEmpPosition.Text = "Position: " & empPosition
        txtRatePerHour.Text = "₱" & hourlyRate.ToString("N2")
        lblPeriod.Text = String.Format("Payroll Period: {0:dd MMM yyyy} - {1:dd MMM yyyy}", startDate, endDate)
        lblTotalHours.Text = totalHours.ToString("0.##")
        lblTotalDays.Text = totalDays
        lblGross.Text = "₱" & gross.ToString("N2")
        lblSSS.Text = "₱" & sss.ToString("N2")
        lblPhilhealth.Text = "₱" & philhealth.ToString("N2")
        lblPagibig.Text = "₱" & pagibig.ToString("N2")
        lblBonus1.Text = "₱" & bonus.ToString("N2")
        lblLateDeduction.Text = "₱" & lateDeduction.ToString("N2")
        lblOtherDeductions.Text = "₱" & otherDeduction.ToString("N2")
        lblTax.Text = "₱" & tax.ToString("N2")
        lblNetPay.Text = "₱" & netPay.ToString("N2")

        ' --- Convert late minutes to HH:MM for print ---
        ' --- Convert late minutes to HH:MM for print ---
        Dim lateHours As Integer = CInt(totalLateMinutes) \ 60
        Dim lateMins As Integer = CInt(totalLateMinutes) Mod 60
        lblLate.Text = lateHours.ToString("0") & ":" & lateMins.ToString("00") ' e.g., 0:16

        ' --- Store globally for printing ---
        totalLateMinutesGlobal = CInt(totalLateMinutes)

        ' --- Save & Print ---
        printingprint()   ' No arguments needed now
        loader()

    End Sub


    ' --- Withholding Tax Function ---
    Private Function ComputeWithholdingTax(gross As Decimal, sss As Decimal, philhealth As Decimal, pagibig As Decimal) As Decimal
        Dim taxable As Decimal = gross - (sss + philhealth + pagibig)

        If taxable <= 16000D Then Return 0D ' No tax for minimum wage earners

        Dim tax As Decimal = 0D
        If taxable <= 33333D Then
            tax = (taxable - 20833D) * 0.2D
        ElseIf taxable <= 66666D Then
            tax = 2500D + (taxable - 33333D) * 0.25D
        ElseIf taxable <= 166666D Then
            tax = 10833.33D + (taxable - 66666D) * 0.3D
        ElseIf taxable <= 666666D Then
            tax = 40833.33D + (taxable - 166666D) * 0.32D
        Else
            tax = 200833.33D + (taxable - 666666D) * 0.35D
        End If
        Return Math.Round(tax, 2)
    End Function
    ' ===========================
    ' Printing Payslip
    ' ===========================
    Private Sub printingprint()
        Dim empID As Integer

        ' --- Determine EmployeeID from DataGridView or ComboBox ---
        If dgvemp.CurrentRow IsNot Nothing Then
            empID = Convert.ToInt32(dgvemp.CurrentRow.Cells("EmployeeID").Value)
        ElseIf cmbEmployee.SelectedIndex <> -1 AndAlso cmbEmployee.SelectedValue IsNot Nothing Then
            If Not Integer.TryParse(cmbEmployee.SelectedValue.ToString(), empID) Then
                MessageBox.Show("Invalid Employee ID.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        Else
            MessageBox.Show("Please select a valid employee.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' --- Get payroll period (adjust if needed) ---
        Dim startDate As Date = Date.Today.AddDays(-15)
        Dim endDate As Date = Date.Today

        ' --- Parse hours and days from labels ---
        Dim totalDays As Integer = Convert.ToInt32(lblTotalDays.Text)
        Dim totalHours As Decimal = Convert.ToDecimal(lblTotalHours.Text)
        lblTotalHours.Text = totalHours.ToString("0.##") ' Format

        ' --- Parse rate per hour ---
        Dim cleanRate As String = txtRatePerHour.Text.Replace("₱", "").Replace(",", "").Trim()
        Dim rate As Decimal
        If Not Decimal.TryParse(cleanRate, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, rate) Then
            MessageBox.Show("Invalid Rate per Hour value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' --- Parse gross and deductions ---
        Dim gross As Decimal = Decimal.Parse(lblGross.Text.Replace("₱", "").Replace(",", "").Trim())
        Dim sss As Decimal = Decimal.Parse(lblSSS.Text.Replace("₱", "").Replace(",", "").Trim())
        Dim philhealth As Decimal = Decimal.Parse(lblPhilhealth.Text.Replace("₱", "").Replace(",", "").Trim())
        Dim pagibig As Decimal = Decimal.Parse(lblPagibig.Text.Replace("₱", "").Replace(",", "").Trim())
        Dim bonus As Decimal = Decimal.Parse(lblBonus1.Text.Replace("₱", "").Replace(",", "").Trim())
        Dim lateDeduction As Decimal = Decimal.Parse(lblLateDeduction.Text.Replace("₱", "").Replace(",", "").Trim())
        Dim otherDeduction As Decimal = Decimal.Parse(lblOtherDeductions.Text.Replace("₱", "").Replace(",", "").Trim())

        ' --- Compute total gross ---
        Dim netGross As Decimal = gross + bonus

        ' --- Compute withholding tax ---
        Dim tax As Decimal = ComputeWithholdingTax(netGross, sss, philhealth, pagibig)

        ' --- Total deduction & net pay ---
        Dim totalDeduction As Decimal = sss + philhealth + pagibig + lateDeduction + otherDeduction + tax
        Dim netPay As Decimal = netGross - totalDeduction

        ' --- Save to database ---
        SavePayslip(empID, startDate, endDate, totalDays, totalHours, rate, netGross, sss, philhealth, pagibig, bonus, lateDeduction, otherDeduction, tax, netPay)

        ' --- Update lblLate using global totalLateMinutesGlobal ---
        Dim lateHours As Integer = totalLateMinutesGlobal \ 60
        Dim lateMins As Integer = totalLateMinutesGlobal Mod 60
        lblLate.Text = lateHours.ToString("0") & ":" & lateMins.ToString("00") ' e.g., 0:16

        ' --- Prepare print ---
        Dim pd As New Printing.PrintDocument()
        AddHandler pd.PrintPage, AddressOf PrintPayslipProfessional

        Dim ppd As New PrintPreviewDialog()
        ppd.Document = pd

        AddHandler ppd.Shown, AddressOf Preview_Shown
        ppd.Show()
    End Sub


    Private Sub SavePayslip(ByVal empID As Integer, ByVal startDate As Date, ByVal endDate As Date,
                            ByVal totalDays As Integer, ByVal totalHours As Decimal,
                            ByVal rate As Decimal, ByVal gross As Decimal, ByVal sss As Decimal,
                            ByVal philhealth As Decimal, ByVal pagibig As Decimal, ByVal bonus As Decimal,
                            ByVal lateDeduction As Decimal, ByVal otherDeduction As Decimal,
                            ByVal tax As Decimal, ByVal netPay As Decimal)

        Dim query As String = "INSERT INTO Payslip " &
                              "(EmployeeID, PeriodStart, PeriodEnd, TotalDaysWorked, TotalHoursWorked, RatePerHour, " &
                              "GrossPay, SSS, PhilHealth, PagIbig, Bonus, LateDeduction, OtherDeduction, Tax, NetPay) " &
                              "VALUES (@empID, @start, @end, @days, @hours, @rate, @gross, @sss, @philhealth, @pagibig, @bonus, @late, @other, @tax, @net)"
        Using conn As New SqlConnection(
            "Data Source=(LocalDB)\MSSQLLocalDB;" &
            "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
            "Integrated Security=True;"
        )

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@empID", empID)
                cmd.Parameters.AddWithValue("@start", startDate)
                cmd.Parameters.AddWithValue("@end", endDate)
                cmd.Parameters.AddWithValue("@days", totalDays)
                cmd.Parameters.AddWithValue("@hours", totalHours)
                cmd.Parameters.AddWithValue("@rate", rate)
                cmd.Parameters.AddWithValue("@gross", gross)
                cmd.Parameters.AddWithValue("@sss", sss)
                cmd.Parameters.AddWithValue("@philhealth", philhealth)
                cmd.Parameters.AddWithValue("@pagibig", pagibig)
                cmd.Parameters.AddWithValue("@bonus", bonus)
                cmd.Parameters.AddWithValue("@late", lateDeduction)
                cmd.Parameters.AddWithValue("@other", otherDeduction)
                cmd.Parameters.AddWithValue("@tax", tax)
                cmd.Parameters.AddWithValue("@net", netPay)

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Private Sub PrintPayslipProfessional(ByVal sender As Object, ByVal e As Printing.PrintPageEventArgs)
        ' --- Fonts ---
        Dim fontTitle As New Font("Arial", 14, FontStyle.Bold)
        Dim fontHeader As New Font("Arial", 12, FontStyle.Bold)
        Dim fontNormal As New Font("Arial", 11)
        Dim fontSmall As New Font("Arial", 9)

        ' --- Layout ---
        Dim pageWidth As Integer = e.PageBounds.Width
        Dim centerX As Integer = pageWidth \ 2
        Dim y As Integer = 50
        Dim contentWidth As Integer = 600
        Dim contentStartX As Integer = centerX - (contentWidth \ 2)
        Dim colWidth As Integer = contentWidth \ 4
        Dim col1X As Integer = contentStartX
        Dim col2X As Integer = contentStartX + colWidth
        Dim col3X As Integer = contentStartX + 2 * colWidth
        Dim col4X As Integer = contentStartX + 3 * colWidth
        Dim sfCenter As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
        Dim rowHeight As Integer = 25

        ' --- Header ---
        e.Graphics.DrawString("WILLTOP HARDWARE AND ELECTRICAL SUPPLY", fontTitle, Brushes.Black, centerX - e.Graphics.MeasureString("WILLTOP HARDWARE AND ELECTRICAL SUPPLY", fontTitle).Width / 2, y)
        y += 30
        e.Graphics.DrawString("Mac Arthur Highway, Dalandanan, Valenzuela City", fontNormal, Brushes.Black, centerX - e.Graphics.MeasureString("Mac Arthur Highway, Dalandanan, Valenzuela City", fontNormal).Width / 2, y)
        y += 20
        e.Graphics.DrawString("Contact: 09995447637", fontNormal, Brushes.Black, centerX - e.Graphics.MeasureString("Contact: 09995447637", fontNormal).Width / 2, y)
        y += 40
        e.Graphics.DrawString("Payroll Period: 01 Sep 2025 - 30 Sep 2025", fontHeader, Brushes.Black, centerX - e.Graphics.MeasureString("Payroll Period: 01 Sep 2025 - 30 Sep 2025", fontHeader).Width / 2, y)
        y += 40

        ' --- Employee Info ---
        e.Graphics.DrawString("Employee ID: " & dgvemp.CurrentRow.Cells("EmployeeID").Value.ToString(), fontNormal, Brushes.Black, col1X, y)
        e.Graphics.DrawString("Position: " & lblEmpPosition.Text.Replace("Position: ", ""), fontNormal, Brushes.Black, col1X, y + 20)
        Dim infoRightX As Integer = col4X + 5
        e.Graphics.DrawString("Name: " & lblEmpName.Text.Replace("Employee Name: ", ""), fontNormal, Brushes.Black, infoRightX, y)
        e.Graphics.DrawString("Rate/Hour: " & txtRatePerHour.Text, fontNormal, Brushes.Black, infoRightX, y + 20)
        y += 60

        ' --- Table Headers ---
        Dim headers() As String = {"Earnings", "Amount", "Deductions", "Amount"}
        For i As Integer = 0 To 4
            e.Graphics.DrawLine(Pens.Black, contentStartX + i * colWidth, y, contentStartX + i * colWidth, y + rowHeight)
        Next
        e.Graphics.DrawLine(Pens.Black, contentStartX, y, contentStartX + contentWidth, y)
        For i As Integer = 0 To 3
            e.Graphics.DrawString(headers(i), fontHeader, Brushes.Black, New Rectangle(contentStartX + i * colWidth, y, colWidth, rowHeight), sfCenter)
        Next
        e.Graphics.DrawLine(Pens.Black, contentStartX, y + rowHeight, contentStartX + contentWidth, y + rowHeight)
        y += rowHeight

        ' --- Earnings & Deductions ---
        Dim earnings() As String = {"Total Hours on Duty", "Total Days on Duty", "Bonus", "Total Late", "Total Income"} ' <-- Only change: "Overtime" → "Total Late"
        Dim deductions() As String = {"SSS", "PhilHealth", "Pag-IBIG", "Late Deduction", "Other Deductions", "Tax"}

        ' Prepare earnings values
        Dim earningsValues(earnings.Length - 1) As String
        earningsValues(0) = lblTotalHours.Text               ' no peso
        earningsValues(1) = lblTotalDays.Text                ' no peso
        earningsValues(2) = ToDecimalSafe(lblBonus1.Text).ToString("N2")   ' bonus
        earningsValues(3) = lblLate.Text ' <-- Shows late hours/minutes
        Dim grossPay As Decimal = ToDecimalSafe(lblGross.Text)
        Dim bonus As Decimal = ToDecimalSafe(lblBonus1.Text)
        Dim totalIncome As Decimal = grossPay + bonus
        earningsValues(4) = totalIncome.ToString("N2")      ' total income

        Dim monetaryEarnings() As Boolean = {False, False, True, False, True} ' Total Late is non-monetary

        ' Prepare deductions values
        Dim deductionsValues(deductions.Length - 1) As String
        deductionsValues(0) = ToDecimalSafe(lblSSS.Text).ToString("N2")
        deductionsValues(1) = ToDecimalSafe(lblPhilhealth.Text).ToString("N2")
        deductionsValues(2) = ToDecimalSafe(lblPagibig.Text).ToString("N2")
        deductionsValues(3) = ToDecimalSafe(lblLateDeduction.Text).ToString("N2")
        deductionsValues(4) = ToDecimalSafe(lblOtherDeductions.Text).ToString("N2")
        deductionsValues(5) = ToDecimalSafe(lblTax.Text).ToString("N2")

        ' --- Draw Table Rows ---
        Dim maxRows As Integer = Math.Max(earnings.Length, deductions.Length)
        For i As Integer = 0 To maxRows - 1
            ' Vertical lines
            For j As Integer = 0 To 4
                e.Graphics.DrawLine(Pens.Black, contentStartX + j * colWidth, y, contentStartX + j * colWidth, y + rowHeight)
            Next
            e.Graphics.DrawLine(Pens.Black, contentStartX, y, contentStartX + contentWidth, y)

            ' Earnings column
            If i < earnings.Length Then e.Graphics.DrawString(earnings(i), fontNormal, Brushes.Black, New Rectangle(col1X, y, colWidth, rowHeight), sfCenter)
            If i < earningsValues.Length Then
                If monetaryEarnings(i) Then
                    e.Graphics.DrawString("₱" & earningsValues(i), fontNormal, Brushes.Black, New Rectangle(col2X, y, colWidth, rowHeight), sfCenter)
                Else
                    e.Graphics.DrawString(earningsValues(i), fontNormal, Brushes.Black, New Rectangle(col2X, y, colWidth, rowHeight), sfCenter)
                End If
            End If

            ' Deductions column
            If i < deductions.Length Then e.Graphics.DrawString(deductions(i), fontNormal, Brushes.Black, New Rectangle(col3X, y, colWidth, rowHeight), sfCenter)
            If i < deductionsValues.Length Then
                e.Graphics.DrawString("₱" & deductionsValues(i), fontNormal, Brushes.Black, New Rectangle(col4X, y, colWidth, rowHeight), sfCenter)
            End If

            ' Bottom line
            e.Graphics.DrawLine(Pens.Black, contentStartX, y + rowHeight, contentStartX + contentWidth, y + rowHeight)
            y += rowHeight
        Next

        ' --- Gross & Net Pay ---
        Dim payRowHeight As Integer = 30
        Dim netPay As Decimal = ToDecimalSafe(lblNetPay.Text)
        For i As Integer = 0 To 4
            e.Graphics.DrawLine(Pens.Black, contentStartX + i * colWidth, y, contentStartX + i * colWidth, y + payRowHeight)
        Next
        e.Graphics.DrawLine(Pens.Black, contentStartX, y, contentStartX + contentWidth, y)
        e.Graphics.DrawString("Gross Pay:", fontHeader, Brushes.Black, New Rectangle(col1X, y, colWidth, payRowHeight), sfCenter)
        e.Graphics.DrawString("₱" & grossPay.ToString("N2"), fontHeader, Brushes.Black, New Rectangle(col2X, y, colWidth, payRowHeight), sfCenter)
        e.Graphics.DrawString("Net Salary:", fontHeader, Brushes.Black, New Rectangle(col3X, y, colWidth, payRowHeight), sfCenter)
        e.Graphics.DrawString("₱" & netPay.ToString("N2"), fontHeader, Brushes.Black, New Rectangle(col4X, y, colWidth, payRowHeight), sfCenter)
        e.Graphics.DrawLine(Pens.Black, contentStartX, y + payRowHeight, contentStartX + contentWidth, y + payRowHeight)
        y += payRowHeight + 40

        ' --- Signatures ---
        Dim sigWidth As Integer = 170
        e.Graphics.DrawLine(Pens.Black, col1X, y, col1X + sigWidth, y)
        e.Graphics.DrawLine(Pens.Black, infoRightX, y, infoRightX + sigWidth, y)
        y += 5
        e.Graphics.DrawString("Employer Signature", fontSmall, Brushes.Black, col1X, y + 10)
        e.Graphics.DrawString("Employee Signature", fontSmall, Brushes.Black, infoRightX, y + 10)
    End Sub

    Private Function GetTotalLateMinutes(empID As Integer, periodStart As Date, periodEnd As Date) As Integer
        Dim totalLate As Integer = 0

        Try
            Using conn As New SqlConnection(
            "Data Source=(LocalDB)\MSSQLLocalDB;" &
            "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
            "Integrated Security=True;"
        )

                conn.Open()

                ' --- Query Attendance within the payroll period ---
                Dim query As String = "SELECT TimeIn, ScheduledIn FROM Attendance " &
                                      "WHERE EmployeeID = @empID AND AttendanceDate BETWEEN @start AND @end"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@empID", empID)
                    cmd.Parameters.AddWithValue("@start", periodStart)
                    cmd.Parameters.AddWithValue("@end", periodEnd)

                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            If Not IsDBNull(reader("TimeIn")) AndAlso Not IsDBNull(reader("ScheduledIn")) Then
                                Dim actual As DateTime = Convert.ToDateTime(reader("TimeIn"))
                                Dim scheduled As DateTime = Convert.ToDateTime(reader("ScheduledIn"))

                                If actual > scheduled Then
                                    totalLate += CInt((actual - scheduled).TotalMinutes) ' add late minutes
                                End If
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error computing total late minutes: " & ex.Message)
        End Try

        Return totalLate
    End Function



    Private Function ToDecimalSafe(input As String) As Decimal
        Dim cleaned As String = input.Replace("₱", "").Replace(",", "").Trim()
        If String.IsNullOrEmpty(cleaned) Then Return 0D
        Dim result As Decimal
        If Decimal.TryParse(cleaned, result) Then
            Return result
        Else
            Return 0D
        End If
    End Function



    Private Sub Guna2GradientButton11_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Guna2GradientButton11.Click
        If dgvemp.CurrentRow Is Nothing Then
            MessageBox.Show("Please select an employee to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim row As DataGridViewRow = dgvemp.CurrentRow
        Dim empId As Integer = Convert.ToInt32(row.Cells("EmployeeID").Value)

        ' Store RatePerHour in a variable first
        Dim rate As Decimal
        Decimal.TryParse(row.Cells("RatePerHour").Value.ToString(), rate) ' safer than Convert.ToDecimal

        ' Create a new instance of Form3
        Dim frm3 As New Form3(
    row.Cells("Name").Value.ToString(),
    row.Cells("Position").Value.ToString(),
    rate,
    row.Cells("Address").Value.ToString(),
    row.Cells("Contact").Value.ToString(),
    row.Cells("Status").Value.ToString()
)

        frm3.Guna2GradientButton2.Text = "UPDATE"
        frm3.Guna2GradientButton2.Tag = empId
        frm3.Show()



    End Sub



    Private Sub lblPeriod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPeriod.Click

    End Sub

    Private Sub Guna2GradientButton5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2GradientButton5.Click
        Dim connectionString As String =
            "Data Source=(LocalDB)\MSSQLLocalDB;" &
            "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
            "Integrated Security=True;"


        If dgvemp.CurrentRow Is Nothing Then
            MsgBox("Please select an employee to delete.", MsgBoxStyle.Exclamation, "No Selection")
            Exit Sub
        End If

        Dim selectedId As Integer = Convert.ToInt32(dgvemp.CurrentRow.Cells("EmployeeID").Value)
        Dim empName As String = dgvemp.CurrentRow.Cells("Name").Value.ToString()


        Using confirm As New Confirmation()
            If confirm.ShowDialog(Me) = DialogResult.OK Then
                If confirm.EnteredPassword = "12345" Then

                    Using myconnection As New SqlConnection(connectionString)
                        Try
                            myconnection.Open()
                            Dim query As String = "DELETE FROM Employee WHERE EmployeeID=@id"
                            Using cmd As New SqlCommand(query, myconnection)
                                cmd.Parameters.AddWithValue("@id", selectedId)
                                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                                If rowsAffected > 0 Then
                                    MsgBox("Employee '" & empName & "' Fired successfully.", MsgBoxStyle.Information, "Success")
                                    viewdata()
                                    LoadEmployeeNames()
                                Else
                                    MsgBox("No employee found with that ID.", MsgBoxStyle.Exclamation)
                                End If
                            End Using
                        Catch ex As Exception
                            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
                        End Try
                    End Using

                Else
                    MessageBox.Show("Invalid admin password.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End Using

        loader()
    End Sub
    Private Sub Guna2rate_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)

        Dim tb As Guna.UI2.WinForms.Guna2TextBox = CType(sender, Guna.UI2.WinForms.Guna2TextBox)


        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        ElseIf Char.IsDigit(e.KeyChar) AndAlso tb.Text.Length >= 3 Then

            e.Handled = True
        End If
    End Sub


    Private Sub Guna2cons_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        ' Allow only digits and control characters (e.g., backspace)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub



    Private Sub Guna2TextBox1_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles Guna2TextBox1.TextChanged
        If Guna2TextBox1.Text.Length > 20 Then
            Guna2TextBox1.Text = Guna2TextBox1.Text.Substring(0, 20)
            Guna2TextBox1.SelectionStart = Guna2TextBox1.Text.Length
        End If
    End Sub

    Private Sub Guna2GradientPanel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Guna2GradientPanel1.Paint

    End Sub

    Private Sub txtRatePerHour_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRatePerHour.Click

    End Sub

    Private Sub lbluser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbluser.Click

    End Sub

    Private Sub Panel1_Paint_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub
    Private Sub LoadCurrentOnDuty1()
        Try
            Using conn As New SqlConnection(
     "Data Source=(LocalDB)\MSSQLLocalDB;" &
     "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
     "Integrated Security=True;"
 )

                conn.Open()

                ' Get employees who clocked in today but haven't clocked out
                Dim query As String = "SELECT EmployeeID, EmployeeName, AttendanceDate, TimeIn, TimeOut " &
                                      "FROM Attendance " &
                                      "WHERE TimeIn IS NOT NULL AND TimeOut IS NULL AND AttendanceDate = CAST(GETDATE() AS DATE)"

                Using da As New SqlDataAdapter(query, conn)
                    Dim dt As New DataTable()
                    da.Fill(dt)



                    todayondutydatagrid.DataSource = dt
                End Using
            End Using

            ' --- Basic grid settings ---
            todayondutydatagrid.ReadOnly = True
            todayondutydatagrid.AllowUserToResizeColumns = False
            todayondutydatagrid.AllowUserToOrderColumns = False
            todayondutydatagrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            todayondutydatagrid.MultiSelect = False
            todayondutydatagrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            ' Hide columns you don't want visible
            If todayondutydatagrid.Columns.Contains("EmployeeID") Then
                todayondutydatagrid.Columns("EmployeeID").Visible = False
            End If
            If todayondutydatagrid.Columns.Contains("AttendanceDate") Then
                todayondutydatagrid.Columns("AttendanceDate").Visible = False
            End If

            ' Optional: highlight late employees
            HighlightLateEmployees(todayondutydatagrid)

            ' Optional: safe styling & adjustments
            StyleAndAdjustDataGrid(todayondutydatagrid, 500)

        Catch ex As Exception
            MsgBox("Error loading current on-duty list: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub



    Private Sub Guna2GradientButton2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2GradientButton2.Click

        Panel1.Visible = False
        Panel2.Visible = False
        cmbEmployee.Text = ""
        Panel6.Visible = True
        bunospanel.Visible = False
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As Object, ByVal e As EventArgs) Handles DateTimePicker1.ValueChanged
        Try
            Dim selectedDate As Date = DateTimePicker1.Value.Date

            Dim connectionString As String =
           "Data Source=(LocalDB)\MSSQLLocalDB;" &
           "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
           "Integrated Security=True;"

            Dim query As String = "SELECT TOP (1000) [PayslipID], [EmployeeID], [PeriodStart], [PeriodEnd], [TotalDaysWorked], " &
                                  "[TotalHoursWorked], [RatePerHour], [GrossPay], [SSS], [PhilHealth], [PagIbig], [NetPay], [DateGenerated] " &
                                  "FROM [Payroll].[dbo].[Payslip] " &
                                  "WHERE CAST(DateGenerated AS DATE) = @selectedDate"

            Using conn As New SqlConnection(connectionString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@selectedDate", selectedDate)
                    Dim dt As New DataTable()
                    Dim adapter As New SqlDataAdapter(cmd)
                    adapter.Fill(dt)

                    If dt.Rows.Count = 0 Then
                        MessageBox.Show("No records found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        DateTimePicker1.Value = DateTime.Now.Date
                        LoadPayslipData()
                    Else
                        c.DataSource = dt
                        StyleAndAdjustDataGrid(c, 400)
                    End If
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error filtering by date: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Guna2GroupBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guna2GroupBox4.Click

    End Sub

    Private Sub Guna2Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Guna2Panel2.Paint

    End Sub

    Private Sub lblDeductionSummary1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblDeductionSummary1.Click

    End Sub


    Private Sub cmbMonth_SelectedIndexChanged_1(sender As System.Object, e As System.EventArgs) Handles cmbMonth.SelectedIndexChanged

    End Sub

    Private Sub Guna2GradientButton4_Click_1(sender As System.Object, e As System.EventArgs) Handles Guna2GradientButton4.Click
        bunospanel.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
        cmbEmployee.Text = ""
        Panel6.Visible = False
    End Sub

    Private Sub SetupDgvBonus()
        dgvBonus.Columns.Clear()
        dgvBonus.AllowUserToAddRows = False ' Important!

        ' EmployeeID column (hidden)
        Dim colEmpID As New DataGridViewTextBoxColumn()
        colEmpID.Name = "EmployeeID"
        colEmpID.HeaderText = "EmployeeID"
        colEmpID.Visible = False
        dgvBonus.Columns.Add(colEmpID)

        ' Employee Name column
        Dim colName As New DataGridViewTextBoxColumn()
        colName.Name = "Name"
        colName.HeaderText = "Employee Name"
        colName.Width = 200
        dgvBonus.Columns.Add(colName)

        dgvBonus.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvBonus.MultiSelect = False
        StyleAndAdjustDataGrid(dgvBonus, 500)
    End Sub
    


    Private Sub Guna2GradientButton8_Click(sender As System.Object, e As System.EventArgs) Handles btnAddToBonus.Click
        If dgvemployee.CurrentRow Is Nothing Then
            MessageBox.Show("Select an employee first.", "No Employee Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Dim empID As Integer = Convert.ToInt32(dgvemployee.CurrentRow.Cells("EmployeeID").Value)
        Dim empName As String = dgvemployee.CurrentRow.Cells("Name").Value.ToString()


        ' Check if employee already exists in dgvBonus
        For Each row As DataGridViewRow In dgvBonus.Rows
            If row.IsNewRow Then Continue For
            If CInt(row.Cells("EmployeeID").Value) = empID Then
                MessageBox.Show("Employee already added to bonus list.", "Duplicate", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        Next


        ' Add employee to dgvBonus
        dgvBonus.Rows.Add(empID, empName)
    End Sub

    Private Sub btnSaveBonus_Click(sender As Object, e As EventArgs) Handles btnSaveBonus.Click
       ' ✅ Confirmation with password before saving
        ' --- Validate bonus amount ---
        Dim bonusAmount As Decimal
        If Not Decimal.TryParse(txtBonus.Text, bonusAmount) Then
            MessageBox.Show("Enter a valid bonus amount.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' --- Validate year input ---
        Dim selectedYear As Integer
        If Not Integer.TryParse(txtYear.Text, selectedYear) Then
            MessageBox.Show("Enter a valid year.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtYear.Focus()
            Exit Sub
        End If

        ' --- Validate month selection ---
        If cmbMonthBonus.SelectedIndex = -1 Then
            MessageBox.Show("Select a month.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Dim selectedMonth As Integer = cmbMonthBonus.SelectedIndex + 1

        ' --- Confirmation before saving ---
        Dim confirm As New Confirmation()
        If confirm.ShowDialog(Me) = DialogResult.OK Then
            If confirm.EnteredPassword = "12345" Then
                ' ✅ Variables bonusAmount, selectedMonth, selectedYear are now accessible here
                Using conn As New SqlConnection(
          "Data Source=(LocalDB)\MSSQLLocalDB;" &
          "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
          "Integrated Security=True;"
      )

                    conn.Open()

                    ' Check duplicates and insert bonuses
                    For Each row As DataGridViewRow In dgvBonus.Rows
                        If row.IsNewRow Then Continue For

                        Dim empID As Integer = CInt(row.Cells("EmployeeID").Value)

                        ' Check duplicate
                        Dim checkQuery As String = "SELECT COUNT(*) FROM EmployeeBonus WHERE EmployeeID=@EmpID AND BonusMonth=@Month AND BonusYear=@Year"
                        Using cmdCheck As New SqlCommand(checkQuery, conn)
                            cmdCheck.Parameters.AddWithValue("@EmpID", empID)
                            cmdCheck.Parameters.AddWithValue("@Month", selectedMonth)
                            cmdCheck.Parameters.AddWithValue("@Year", selectedYear)
                            Dim exists As Integer = CInt(cmdCheck.ExecuteScalar())
                            If exists > 0 Then
                                MessageBox.Show("Bonus already exists for " & row.Cells("Name").Value.ToString() & " in " & cmbMonthBonus.Text & ".", "Duplicate Bonus", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                Exit Sub
                            End If
                        End Using

                        ' Insert bonus
                        Dim insertQuery As String = "INSERT INTO EmployeeBonus (EmployeeID, BonusAmount, BonusMonth, BonusYear) VALUES (@EmpID, @Amount, @Month, @Year)"
                        Using cmdInsert As New SqlCommand(insertQuery, conn)
                            cmdInsert.Parameters.AddWithValue("@EmpID", empID)
                            cmdInsert.Parameters.AddWithValue("@Amount", bonusAmount)
                            cmdInsert.Parameters.AddWithValue("@Month", selectedMonth)
                            cmdInsert.Parameters.AddWithValue("@Year", selectedYear)
                            cmdInsert.ExecuteNonQuery()
                        End Using
                    Next
                End Using

                dgvBonus.Rows.Clear()
                LoadEmployeeBonuses()
                MessageBox.Show("Bonus saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Invalid admin password.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        Else
            Exit Sub
        End If


    End Sub


  
    Private Sub SetupCmbMonthBonus()
        cmbMonthBonus.Items.Clear()

        ' Add months
        cmbMonthBonus.Items.Add("January")
        cmbMonthBonus.Items.Add("February")
        cmbMonthBonus.Items.Add("March")
        cmbMonthBonus.Items.Add("April")
        cmbMonthBonus.Items.Add("May")
        cmbMonthBonus.Items.Add("June")
        cmbMonthBonus.Items.Add("July")
        cmbMonthBonus.Items.Add("August")
        cmbMonthBonus.Items.Add("September")
        cmbMonthBonus.Items.Add("October")
        cmbMonthBonus.Items.Add("November")
        cmbMonthBonus.Items.Add("December")

        ' Optional: select current month by default
        cmbMonthBonus.SelectedIndex = DateTime.Now.Month - 1
    End Sub

    Private Sub cmbMonthbunos_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbMonthBonus.SelectedIndexChanged

    End Sub

    Private Sub btnREMOVEToBonus_Click(sender As System.Object, e As System.EventArgs) Handles btnREMOVEToBonus.Click
        If dgvBonus.CurrentRow Is Nothing OrElse dgvBonus.CurrentRow.IsNewRow Then
            MessageBox.Show("Select an employee from the bonus list to remove.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Confirm deletion
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to remove " & dgvBonus.CurrentRow.Cells("Name").Value.ToString() & " from the bonus list?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            dgvBonus.Rows.Remove(dgvBonus.CurrentRow)
        End If
    End Sub
    Private Sub datetimepickerbunossorter_ValueChanged(sender As Object, e As EventArgs) Handles datetimepickerbunossorter.ValueChanged
        ' Use only month/year, ignore day
        Dim filterDate As New DateTime(datetimepickerbunossorter.Value.Year, datetimepickerbunossorter.Value.Month, 1)
        LoadEmployeeBonuses(filterDate)
    End Sub

    Public Sub LoadEmployeeBonuses(Optional ByVal filterDate As DateTime? = Nothing)
        Using conn As New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database\payroll.mdf;Integrated Security=True;")

            Dim query As String = "SELECT b.BonusID, e.Name AS EmployeeName, b.BonusAmount, b.BonusMonth, b.BonusYear, b.DateAdded " &
                                  "FROM EmployeeBonus b " &
                                  "INNER JOIN Employee e ON b.EmployeeID = e.EmployeeID "

            ' Filter by selected month/year
            If filterDate.HasValue Then
                query &= "WHERE b.BonusMonth = @Month AND b.BonusYear = @Year "
            End If

            query &= "ORDER BY b.BonusYear DESC, b.BonusMonth DESC"

            Dim da As New SqlDataAdapter(query, conn)
            If filterDate.HasValue Then
                da.SelectCommand.Parameters.AddWithValue("@Month", filterDate.Value.Month)
                da.SelectCommand.Parameters.AddWithValue("@Year", filterDate.Value.Year)
            End If

            Dim dt As New DataTable()
            da.Fill(dt)

            dgvEmployeeBonus.DataSource = dt
            StyleAndAdjustDataGrid(dgvEmployeeBonus, 500)

            ' Hide BonusID
            If dgvEmployeeBonus.Columns.Contains("BonusID") Then dgvEmployeeBonus.Columns("BonusID").Visible = False

            ' Format BonusAmount as peso
            Dim pesoFormat As New CultureInfo("en-PH", False)
            pesoFormat.NumberFormat.CurrencySymbol = "₱"
            dgvEmployeeBonus.Columns("BonusAmount").DefaultCellStyle.Format = "C2"
            dgvEmployeeBonus.Columns("BonusAmount").DefaultCellStyle.FormatProvider = pesoFormat

            dgvEmployeeBonus.AllowUserToAddRows = False
            dgvEmployeeBonus.ReadOnly = True
        End Using
    End Sub

    Private Sub Guna2GradientButton8_Click_1(sender As System.Object, e As System.EventArgs) Handles Guna2GradientButton8.Click
         Dim currentMonthYear As New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        datetimepickerbunossorter.Value = currentMonthYear

        ' Reload bonuses for the current month/year
        LoadEmployeeBonuses(currentMonthYear)
    End Sub

    Private Sub bunospanel_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles bunospanel.Paint

    End Sub

    Private Sub Guna2Panel9_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles Guna2Panel9.Paint

    End Sub

    Private Sub Guna2Panel11_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles Guna2Panel11.Paint

    End Sub
End Class