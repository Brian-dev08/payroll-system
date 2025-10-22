CREATE TABLE Attendance (
    ID INT IDENTITY(1,1) PRIMARY KEY,
    EmployeeID INT,
    AttendanceDate DATE,
    TimeIn TIME,
    TimeOut TIME,
	 EmployeeName NVARCHAR(100)
);
