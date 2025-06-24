' =====================================================
' CYBERLINK PUBLISHING COMPANY MANAGEMENT SYSTEM
' Complete Excel VBA Solution
' =====================================================

Option Explicit

' Global Variables
Public Const COMPANY_NAME As String = "Cyberlink Publishing Company"
Public Const VERSION As String = "1.0"

' Worksheet Names
Public Const WS_DASHBOARD As String = "Dashboard"
Public Const WS_CUSTOMERS As String = "Customers"
Public Const WS_ORDERS As String = "Orders"
Public Const WS_EMPLOYEES As String = "Employees"
Public Const WS_PAYMENTS As String = "Payments"
Public Const WS_REPORTS As String = "Reports"
Public Const WS_SETTINGS As String = "Settings"

' =====================================================
' MAIN INITIALIZATION MODULE
' =====================================================

Sub InitializeSystem()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Create all worksheets
    Call CreateWorksheets
    
    ' Setup all data structures
    Call SetupCustomersSheet
    Call SetupOrdersSheet
    Call SetupEmployeesSheet
    Call SetupPaymentsSheet
    Call SetupDashboard
    Call SetupReportsSheet
    Call SetupSettingsSheet
    
    ' Generate sample data
    Call GenerateSampleData
    
    ' Create navigation menu
    Call CreateNavigationMenu
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Cyberlink Publishing Management System initialized successfully!", vbInformation, COMPANY_NAME
End Sub

' =====================================================
' WORKSHEET CREATION MODULE
' =====================================================

Sub CreateWorksheets()
    Dim ws As Worksheet
    Dim wsNames As Variant
    Dim i As Integer
    
    wsNames = Array(WS_DASHBOARD, WS_CUSTOMERS, WS_ORDERS, WS_EMPLOYEES, WS_PAYMENTS, WS_REPORTS, WS_SETTINGS)
    
    ' Delete existing sheets (except first one)
    For i = ActiveWorkbook.Worksheets.Count To 2 Step -1
        ActiveWorkbook.Worksheets(i).Delete
    Next i
    
    ' Rename first sheet to Dashboard
    ActiveWorkbook.Worksheets(1).Name = WS_DASHBOARD
    
    ' Create remaining sheets
    For i = 1 To UBound(wsNames)
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = wsNames(i)
    Next i
End Sub

' =====================================================
' CUSTOMERS MODULE
' =====================================================

Sub SetupCustomersSheet()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(WS_CUSTOMERS)
    
    With ws
        .Cells.Clear
        
        ' Header styling
        .Range("A1:J1").Interior.Color = RGB(37, 99, 235)
        .Range("A1:J1").Font.Color = RGB(255, 255, 255)
        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").Font.Size = 12
        
        ' Headers
        .Cells(1, 1) = "Customer ID"
        .Cells(1, 2) = "Name"
        .Cells(1, 3) = "Account Number"
        .Cells(1, 4) = "Company Name"
        .Cells(1, 5) = "Contact Number"
        .Cells(1, 6) = "Email"
        .Cells(1, 7) = "Created Date"
        .Cells(1, 8) = "Total Orders"
        .Cells(1, 9) = "Total Spent"
        .Cells(1, 10) = "Status"
        
        ' Column widths
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:B").ColumnWidth = 20
        .Columns("C:C").ColumnWidth = 15
        .Columns("D:D").ColumnWidth = 25
        .Columns("E:E").ColumnWidth = 15
        .Columns("F:F").ColumnWidth = 25
        .Columns("G:G").ColumnWidth = 12
        .Columns("H:H").ColumnWidth = 12
        .Columns("I:I").ColumnWidth = 15
        .Columns("J:J").ColumnWidth = 10
        
        ' Freeze panes
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
End Sub

Sub AddCustomer()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim customerID As String
    Dim customerName As String
    Dim companyName As String
    Dim contactNumber As String
    Dim email As String
    
    Set ws = ActiveWorkbook.Worksheets(WS_CUSTOMERS)
    
    ' Get input from user
    customerName = InputBox("Enter Customer Name:", "Add Customer")
    If customerName = "" Then Exit Sub
    
    companyName = InputBox("Enter Company Name:", "Add Customer")
    If companyName = "" Then Exit Sub
    
    contactNumber = InputBox("Enter Contact Number:", "Add Customer")
    If contactNumber = "" Then Exit Sub
    
    email = InputBox("Enter Email Address:", "Add Customer")
    If email = "" Then Exit Sub
    
    ' Generate customer ID
    customerID = "CUS-" & Format(Now, "yyyymmddhhmmss")
    
    ' Find next row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Add customer data
    With ws
        .Cells(lastRow, 1) = customerID
        .Cells(lastRow, 2) = customerName
        .Cells(lastRow, 3) = "ACC-" & Right(customerID, 6)
        .Cells(lastRow, 4) = companyName
        .Cells(lastRow, 5) = contactNumber
        .Cells(lastRow, 6) = email
        .Cells(lastRow, 7) = Date
        .Cells(lastRow, 8) = 0
        .Cells(lastRow, 9) = 0
        .Cells(lastRow, 10) = "Active"
    End With
    
    MsgBox "Customer added successfully!", vbInformation
    Call UpdateDashboard
End Sub

Sub EditCustomer()
    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim customerID As String
    
    Set ws = ActiveWorkbook.Worksheets(WS_CUSTOMERS)
    selectedRow = ActiveCell.Row
    
    If selectedRow <= 1 Then
        MsgBox "Please select a customer row to edit.", vbWarning
        Exit Sub
    End If
    
    customerID = ws.Cells(selectedRow, 1).Value
    
    ' Edit customer details
    ws.Cells(selectedRow, 2) = InputBox("Customer Name:", "Edit Customer", ws.Cells(selectedRow, 2))
    ws.Cells(selectedRow, 4) = InputBox("Company Name:", "Edit Customer", ws.Cells(selectedRow, 4))
    ws.Cells(selectedRow, 5) = InputBox("Contact Number:", "Edit Customer", ws.Cells(selectedRow, 5))
    ws.Cells(selectedRow, 6) = InputBox("Email:", "Edit Customer", ws.Cells(selectedRow, 6))
    
    MsgBox "Customer updated successfully!", vbInformation
End Sub

Sub DeleteCustomer()
    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim response As VbMsgBoxResult
    
    Set ws = ActiveWorkbook.Worksheets(WS_CUSTOMERS)
    selectedRow = ActiveCell.Row
    
    If selectedRow <= 1 Then
        MsgBox "Please select a customer row to delete.", vbWarning
        Exit Sub
    End If
    
    response = MsgBox("Are you sure you want to delete this customer?", vbYesNo + vbQuestion)
    If response = vbYes Then
        ws.Rows(selectedRow).Delete
        MsgBox "Customer deleted successfully!", vbInformation
        Call UpdateDashboard
    End If
End Sub

' =====================================================
' ORDERS MODULE
' =====================================================

Sub SetupOrdersSheet()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(WS_ORDERS)
    
    With ws
        .Cells.Clear
        
        ' Header styling
        .Range("A1:P1").Interior.Color = RGB(37, 99, 235)
        .Range("A1:P1").Font.Color = RGB(255, 255, 255)
        .Range("A1:P1").Font.Bold = True
        .Range("A1:P1").Font.Size = 12
        
        ' Headers
        .Cells(1, 1) = "Order ID"
        .Cells(1, 2) = "Customer ID"
        .Cells(1, 3) = "Customer Name"
        .Cells(1, 4) = "Order Type"
        .Cells(1, 5) = "Description"
        .Cells(1, 6) = "Quantity"
        .Cells(1, 7) = "Unit Price"
        .Cells(1, 8) = "Total Price"
        .Cells(1, 9) = "Order Date"
        .Cells(1, 10) = "Status"
        .Cells(1, 11) = "Payment Status"
        .Cells(1, 12) = "Assigned Employee"
        .Cells(1, 13) = "Estimated Delivery"
        .Cells(1, 14) = "Actual Delivery"
        .Cells(1, 15) = "Priority"
        .Cells(1, 16) = "Notes"
        
        ' Column widths
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:B").ColumnWidth = 15
        .Columns("C:C").ColumnWidth = 20
        .Columns("D:D").ColumnWidth = 18
        .Columns("E:E").ColumnWidth = 30
        .Columns("F:F").ColumnWidth = 10
        .Columns("G:G").ColumnWidth = 12
        .Columns("H:H").ColumnWidth = 12
        .Columns("I:I").ColumnWidth = 12
        .Columns("J:J").ColumnWidth = 15
        .Columns("K:K").ColumnWidth = 15
        .Columns("L:L").ColumnWidth = 20
        .Columns("M:M").ColumnWidth = 15
        .Columns("N:N").ColumnWidth = 15
        .Columns("O:O").ColumnWidth = 10
        .Columns("P:P").ColumnWidth = 30
        
        ' Freeze panes
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
End Sub

Sub AddOrder()
    Dim ws As Worksheet
    Dim wsCustomers As Worksheet
    Dim wsEmployees As Worksheet
    Dim lastRow As Long
    Dim orderID As String
    Dim customerList As String
    Dim employeeList As String
    Dim i As Long
    
    Set ws = ActiveWorkbook.Worksheets(WS_ORDERS)
    Set wsCustomers = ActiveWorkbook.Worksheets(WS_CUSTOMERS)
    Set wsEmployees = ActiveWorkbook.Worksheets(WS_EMPLOYEES)
    
    ' Build customer list
    customerList = ""
    For i = 2 To wsCustomers.Cells(wsCustomers.Rows.Count, 1).End(xlUp).Row
        customerList = customerList & wsCustomers.Cells(i, 1) & " - " & wsCustomers.Cells(i, 2) & vbCrLf
    Next i
    
    ' Build employee list
    employeeList = ""
    For i = 2 To wsEmployees.Cells(wsEmployees.Rows.Count, 1).End(xlUp).Row
        employeeList = employeeList & wsEmployees.Cells(i, 1) & " - " & wsEmployees.Cells(i, 2) & vbCrLf
    Next i
    
    ' Show order form
    Call ShowOrderForm("", "ADD")
End Sub

Sub ShowOrderForm(orderID As String, mode As String)
    ' This would typically show a UserForm for order entry
    ' For simplicity, using InputBox method
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim customerID As String
    Dim orderType As String
    Dim description As String
    Dim quantity As Long
    Dim unitPrice As Double
    Dim employeeID As String
    
    Set ws = ActiveWorkbook.Worksheets(WS_ORDERS)
    
    customerID = InputBox("Enter Customer ID:", "Order Form")
    If customerID = "" Then Exit Sub
    
    orderType = InputBox("Enter Order Type (Book Publishing, Magazine Layout, etc.):", "Order Form")
    If orderType = "" Then Exit Sub
    
    description = InputBox("Enter Description:", "Order Form")
    If description = "" Then Exit Sub
    
    quantity = Val(InputBox("Enter Quantity:", "Order Form", "1"))
    unitPrice = Val(InputBox("Enter Unit Price:", "Order Form", "100"))
    
    employeeID = InputBox("Enter Employee ID:", "Order Form")
    If employeeID = "" Then Exit Sub
    
    If mode = "ADD" Then
        orderID = "ORD-" & Format(Now, "yyyymmddhhmmss")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        
        With ws
            .Cells(lastRow, 1) = orderID
            .Cells(lastRow, 2) = customerID
            .Cells(lastRow, 3) = GetCustomerName(customerID)
            .Cells(lastRow, 4) = orderType
            .Cells(lastRow, 5) = description
            .Cells(lastRow, 6) = quantity
            .Cells(lastRow, 7) = unitPrice
            .Cells(lastRow, 8) = quantity * unitPrice
            .Cells(lastRow, 9) = Date
            .Cells(lastRow, 10) = "Ordered"
            .Cells(lastRow, 11) = "Bid"
            .Cells(lastRow, 12) = GetEmployeeName(employeeID)
            .Cells(lastRow, 13) = Date + 30
            .Cells(lastRow, 14) = ""
            .Cells(lastRow, 15) = "Medium"
            .Cells(lastRow, 16) = ""
        End With
        
        ' Create corresponding payment record
        Call CreatePaymentRecord(orderID, customerID, quantity * unitPrice)
        
        MsgBox "Order created successfully!", vbInformation
    End If
    
    Call UpdateDashboard
End Sub

Function GetCustomerName(customerID As String) As String
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ActiveWorkbook.Worksheets(WS_CUSTOMERS)
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1) = customerID Then
            GetCustomerName = ws.Cells(i, 2)
            Exit Function
        End If
    Next i
    
    GetCustomerName = "Unknown"
End Function

Function GetEmployeeName(employeeID As String) As String
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ActiveWorkbook.Worksheets(WS_EMPLOYEES)
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1) = employeeID Then
            GetEmployeeName = ws.Cells(i, 2)
            Exit Function
        End If
    Next i
    
    GetEmployeeName = "Unknown"
End Function

Sub UpdateOrderStatus()
    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim newStatus As String
    
    Set ws = ActiveWorkbook.Worksheets(WS_ORDERS)
    selectedRow = ActiveCell.Row
    
    If selectedRow <= 1 Then
        MsgBox "Please select an order row to update.", vbWarning
        Exit Sub
    End If
    
    newStatus = InputBox("Enter new status (Ordered, In Progress, On Schedule, Completed, Shipped):", "Update Status", ws.Cells(selectedRow, 10))
    
    If newStatus <> "" Then
        ws.Cells(selectedRow, 10) = newStatus
        
        ' If completed or shipped, set actual delivery date
        If newStatus = "Completed" Or newStatus = "Shipped" Then
            ws.Cells(selectedRow, 14) = Date
        End If
        
        MsgBox "Order status updated successfully!", vbInformation
        Call UpdateDashboard
    End If
End Sub

' =====================================================
' EMPLOYEES MODULE
' =====================================================

Sub SetupEmployeesSheet()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(WS_EMPLOYEES)
    
    With ws
        .Cells.Clear
        
        ' Header styling
        .Range("A1:L1").Interior.Color = RGB(37, 99, 235)
        .Range("A1:L1").Font.Color = RGB(255, 255, 255)
        .Range("A1:L1").Font.Bold = True
        .Range("A1:L1").Font.Size = 12
        
        ' Headers
        .Cells(1, 1) = "Employee ID"
        .Cells(1, 2) = "Name"
        .Cells(1, 3) = "Role"
        .Cells(1, 4) = "Responsibility"
        .Cells(1, 5) = "Email"
        .Cells(1, 6) = "Phone"
        .Cells(1, 7) = "Join Date"
        .Cells(1, 8) = "Skill Level"
        .Cells(1, 9) = "Assigned Orders"
        .Cells(1, 10) = "Completed Tasks"
        .Cells(1, 11) = "Performance Rating"
        .Cells(1, 12) = "Status"
        
        ' Column widths
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:B").ColumnWidth = 20
        .Columns("C:C").ColumnWidth = 20
        .Columns("D:D").ColumnWidth = 30
        .Columns("E:E").ColumnWidth = 25
        .Columns("F:F").ColumnWidth = 15
        .Columns("G:G").ColumnWidth = 12
        .Columns("H:H").ColumnWidth = 12
        .Columns("I:I").ColumnWidth = 15
        .Columns("J:J").ColumnWidth = 15
        .Columns("K:K").ColumnWidth = 15
        .Columns("L:L").ColumnWidth = 10
        
        ' Freeze panes
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
End Sub

Sub AddEmployee()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim employeeID As String
    Dim employeeName As String
    Dim role As String
    Dim responsibility As String
    Dim email As String
    Dim phone As String
    Dim skillLevel As String
    
    Set ws = ActiveWorkbook.Worksheets(WS_EMPLOYEES)
    
    ' Get input from user
    employeeName = InputBox("Enter Employee Name:", "Add Employee")
    If employeeName = "" Then Exit Sub
    
    role = InputBox("Enter Role:", "Add Employee")
    If role = "" Then Exit Sub
    
    responsibility = InputBox("Enter Responsibility:", "Add Employee")
    If responsibility = "" Then Exit Sub
    
    email = InputBox("Enter Email:", "Add Employee")
    If email = "" Then Exit Sub
    
    phone = InputBox("Enter Phone:", "Add Employee")
    If phone = "" Then Exit Sub
    
    skillLevel = InputBox("Enter Skill Level (Junior, Intermediate, Senior, Expert):", "Add Employee", "Intermediate")
    
    ' Generate employee ID
    employeeID = "EMP-" & Format(Now, "yyyymmddhhmmss")
    
    ' Find next row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Add employee data
    With ws
        .Cells(lastRow, 1) = employeeID
        .Cells(lastRow, 2) = employeeName
        .Cells(lastRow, 3) = role
        .Cells(lastRow, 4) = responsibility
        .Cells(lastRow, 5) = email
        .Cells(lastRow, 6) = phone
        .Cells(lastRow, 7) = Date
        .Cells(lastRow, 8) = skillLevel
        .Cells(lastRow, 9) = 0
        .Cells(lastRow, 10) = 0
        .Cells(lastRow, 11) = 4.0
        .Cells(lastRow, 12) = "Active"
    End With
    
    MsgBox "Employee added successfully!", vbInformation
    Call UpdateDashboard
End Sub

' =====================================================
' PAYMENTS MODULE
' =====================================================

Sub SetupPaymentsSheet()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(WS_PAYMENTS)
    
    With ws
        .Cells.Clear
        
        ' Header styling
        .Range("A1:L1").Interior.Color = RGB(37, 99, 235)
        .Range("A1:L1").Font.Color = RGB(255, 255, 255)
        .Range("A1:L1").Font.Bold = True
        .Range("A1:L1").Font.Size = 12
        
        ' Headers
        .Cells(1, 1) = "Payment ID"
        .Cells(1, 2) = "Order ID"
        .Cells(1, 3) = "Customer ID"
        .Cells(1, 4) = "Total Amount"
        .Cells(1, 5) = "Paid Amount"
        .Cells(1, 6) = "Outstanding Amount"
        .Cells(1, 7) = "Payment Date"
        .Cells(1, 8) = "Due Date"
        .Cells(1, 9) = "Status"
        .Cells(1, 10) = "Payment Method"
        .Cells(1, 11) = "Transaction ID"
        .Cells(1, 12) = "Notes"
        
        ' Column widths
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:B").ColumnWidth = 15
        .Columns("C:C").ColumnWidth = 15
        .Columns("D:D").ColumnWidth = 15
        .Columns("E:E").ColumnWidth = 15
        .Columns("F:F").ColumnWidth = 15
        .Columns("G:G").ColumnWidth = 12
        .Columns("H:H").ColumnWidth = 12
        .Columns("I:I").ColumnWidth = 15
        .Columns("J:J").ColumnWidth = 15
        .Columns("K:K").ColumnWidth = 15
        .Columns("L:L").ColumnWidth = 30
        
        ' Freeze panes
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
End Sub

Sub CreatePaymentRecord(orderID As String, customerID As String, amount As Double)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim paymentID As String
    
    Set ws = ActiveWorkbook.Worksheets(WS_PAYMENTS)
    
    paymentID = "PAY-" & Format(Now, "yyyymmddhhmmss")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    With ws
        .Cells(lastRow, 1) = paymentID
        .Cells(lastRow, 2) = orderID
        .Cells(lastRow, 3) = customerID
        .Cells(lastRow, 4) = amount
        .Cells(lastRow, 5) = 0
        .Cells(lastRow, 6) = amount
        .Cells(lastRow, 7) = ""
        .Cells(lastRow, 8) = Date + 30
        .Cells(lastRow, 9) = "Unpaid"
        .Cells(lastRow, 10) = ""
        .Cells(lastRow, 11) = ""
        .Cells(lastRow, 12) = ""
    End With
End Sub

Sub RecordPayment()
    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim paymentAmount As Double
    Dim paymentMethod As String
    Dim transactionID As String
    
    Set ws = ActiveWorkbook.Worksheets(WS_PAYMENTS)
    selectedRow = ActiveCell.Row
    
    If selectedRow <= 1 Then
        MsgBox "Please select a payment row to update.", vbWarning
        Exit Sub
    End If
    
    paymentAmount = Val(InputBox("Enter payment amount:", "Record Payment"))
    If paymentAmount <= 0 Then Exit Sub
    
    paymentMethod = InputBox("Enter payment method:", "Record Payment", "Credit Card")
    transactionID = InputBox("Enter transaction ID:", "Record Payment")
    
    With ws
        .Cells(selectedRow, 5) = .Cells(selectedRow, 5) + paymentAmount
        .Cells(selectedRow, 6) = .Cells(selectedRow, 4) - .Cells(selectedRow, 5)
        .Cells(selectedRow, 7) = Date
        .Cells(selectedRow, 10) = paymentMethod
        .Cells(selectedRow, 11) = transactionID
        
        If .Cells(selectedRow, 6) <= 0 Then
            .Cells(selectedRow, 9) = "Paid Full"
        Else
            .Cells(selectedRow, 9) = "Partial Payment"
        End If
    End With
    
    MsgBox "Payment recorded successfully!", vbInformation
    Call UpdateDashboard
End Sub

' =====================================================
' DASHBOARD MODULE
' =====================================================

Sub SetupDashboard()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(WS_DASHBOARD)
    
    With ws
        .Cells.Clear
        
        ' Title
        .Range("A1:J1").Merge
        .Range("A1").Value = COMPANY_NAME & " - Management Dashboard"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(37, 99, 235)
        .Range("A1").Font.Color = RGB(255, 255, 255)
        
        ' KPI Headers
        .Range("A3").Value = "KEY PERFORMANCE INDICATORS"
        .Range("A3").Font.Size = 14
        .Range("A3").Font.Bold = True
        
        ' KPI Labels
        .Range("A5").Value = "Total Customers:"
        .Range("A6").Value = "Total Orders:"
        .Range("A7").Value = "Total Revenue:"
        .Range("A8").Value = "Pending Payments:"
        .Range("A9").Value = "Active Employees:"
        .Range("A10").Value = "Completed Orders:"
        
        ' Recent Activity Header
        .Range("A12").Value = "RECENT ACTIVITY"
        .Range("A12").Font.Size = 14
        .Range("A12").Font.Bold = True
        
        ' Activity Headers
        .Range("A14:F14").Interior.Color = RGB(240, 240, 240)
        .Range("A14").Value = "Date"
        .Range("B14").Value = "Type"
        .Range("C14").Value = "Description"
        .Range("D14").Value = "Amount"
        .Range("E14").Value = "Status"
        .Range("F14").Value = "Employee"
        
        ' Column widths
        .Columns("A:A").ColumnWidth = 12
        .Columns("B:B").ColumnWidth = 15
        .Columns("C:C").ColumnWidth = 40
        .Columns("D:D").ColumnWidth = 15
        .Columns("E:E").ColumnWidth = 15
        .Columns("F:F").ColumnWidth = 20
    End With
    
    Call UpdateDashboard
End Sub

Sub UpdateDashboard()
    Dim ws As Worksheet
    Dim wsCustomers As Worksheet
    Dim wsOrders As Worksheet
    Dim wsPayments As Worksheet
    Dim wsEmployees As Worksheet
    
    Set ws = ActiveWorkbook.Worksheets(WS_DASHBOARD)
    Set wsCustomers = ActiveWorkbook.Worksheets(WS_CUSTOMERS)
    Set wsOrders = ActiveWorkbook.Worksheets(WS_ORDERS)
    Set wsPayments = ActiveWorkbook.Worksheets(WS_PAYMENTS)
    Set wsEmployees = ActiveWorkbook.Worksheets(WS_EMPLOYEES)
    
    ' Update KPIs
    ws.Range("B5").Value = CountRows(wsCustomers) - 1
    ws.Range("B6").Value = CountRows(wsOrders) - 1
    ws.Range("B7").Value = Format(CalculateTotalRevenue(), "Currency")
    ws.Range("B8").Value = CountPendingPayments()
    ws.Range("B9").Value = CountRows(wsEmployees) - 1
    ws.Range("B10").Value = CountCompletedOrders()
    
    ' Update timestamp
    ws.Range("H1").Value = "Last Updated: " & Format(Now, "mm/dd/yyyy hh:mm")
End Sub

Function CountRows(ws As Worksheet) As Long
    CountRows = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function

Function CalculateTotalRevenue() As Double
    Dim ws As Worksheet
    Dim i As Long
    Dim total As Double
    
    Set ws = ActiveWorkbook.Worksheets(WS_PAYMENTS)
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        total = total + ws.Cells(i, 5).Value ' Paid Amount column
    Next i
    
    CalculateTotalRevenue = total
End Function

Function CountPendingPayments() As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim count As Long
    
    Set ws = ActiveWorkbook.Worksheets(WS_PAYMENTS)
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 6).Value > 0 Then ' Outstanding Amount > 0
            count = count + 1
        End If
    Next i
    
    CountPendingPayments = count
End Function

Function CountCompletedOrders() As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim count As Long
    
    Set ws = ActiveWorkbook.Worksheets(WS_ORDERS)
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 10).Value = "Completed" Or ws.Cells(i, 10).Value = "Shipped" Then
            count = count + 1
        End If
    Next i
    
    CountCompletedOrders = count
End Function

' =====================================================
' REPORTS MODULE
' =====================================================

Sub SetupReportsSheet()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(WS_REPORTS)
    
    With ws
        .Cells.Clear
        
        ' Title
        .Range("A1").Value = "REPORTS & ANALYTICS"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        ' Report buttons (would be actual buttons in full implementation)
        .Range("A3").Value = "Available Reports:"
        .Range("A5").Value = "1. Customer Report"
        .Range("A6").Value = "2. Orders Report"
        .Range("A7").Value = "3. Payments Report"
        .Range("A8").Value = "4. Employee Performance Report"
        .Range("A9").Value = "5. Financial Summary Report"
        .Range("A10").Value = "6. Monthly Revenue Report"
        
        .Range("A12").Value = "Export Options:"
        .Range("A14").Value = "• Export to Excel"
        .Range("A15").Value = "• Export to PDF"
        .Range("A16").Value = "• Export to CSV"
    End With
End Sub

Sub GenerateCustomerReport()
    Dim ws As Worksheet
    Dim wsReport As Worksheet
    Dim i As Long
    Dim lastRow As Long
    
    Set ws = ActiveWorkbook.Worksheets(WS_CUSTOMERS)
    Set wsReport = ActiveWorkbook.Worksheets.Add
    wsReport.Name = "Customer_Report_" & Format(Date, "yyyymmdd")
    
    ' Copy headers
    ws.Range("A1:J1").Copy wsReport.Range("A1")
    
    ' Copy data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        ws.Range("A2:J" & lastRow).Copy wsReport.Range("A2")
    End If
    
    ' Format report
    With wsReport
        .Range("A1:J1").Interior.Color = RGB(37, 99, 235)
        .Range("A1:J1").Font.Color = RGB(255, 255, 255)
        .Range("A1:J1").Font.Bold = True
        .Columns.AutoFit
    End With
    
    MsgBox "Customer report generated successfully!", vbInformation
End Sub

Sub GenerateOrdersReport()
    Dim ws As Worksheet
    Dim wsReport As Worksheet
    Dim i As Long
    Dim lastRow As Long
    
    Set ws = ActiveWorkbook.Worksheets(WS_ORDERS)
    Set wsReport = ActiveWorkbook.Worksheets.Add
    wsReport.Name = "Orders_Report_" & Format(Date, "yyyymmdd")
    
    ' Copy headers
    ws.Range("A1:P1").Copy wsReport.Range("A1")
    
    ' Copy data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        ws.Range("A2:P" & lastRow).Copy wsReport.Range("A2")
    End If
    
    ' Format report
    With wsReport
        .Range("A1:P1").Interior.Color = RGB(37, 99, 235)
        .Range("A1:P1").Font.Color = RGB(255, 255, 255)
        .Range("A1:P1").Font.Bold = True
        .Columns.AutoFit
    End With
    
    MsgBox "Orders report generated successfully!", vbInformation
End Sub

Sub GeneratePaymentsReport()
    Dim ws As Worksheet
    Dim wsReport As Worksheet
    Dim i As Long
    Dim lastRow As Long
    
    Set ws = ActiveWorkbook.Worksheets(WS_PAYMENTS)
    Set wsReport = ActiveWorkbook.Worksheets.Add
    wsReport.Name = "Payments_Report_" & Format(Date, "yyyymmdd")
    
    ' Copy headers
    ws.Range("A1:L1").Copy wsReport.Range("A1")
    
    ' Copy data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        ws.Range("A2:L" & lastRow).Copy wsReport.Range("A2")
    End If
    
    ' Format report
    With wsReport
        .Range("A1:L1").Interior.Color = RGB(37, 99, 235)
        .Range("A1:L1").Font.Color = RGB(255, 255, 255)
        .Range("A1:L1").Font.Bold = True
        .Columns.AutoFit
    End With
    
    MsgBox "Payments report generated successfully!", vbInformation
End Sub

Sub GenerateFinancialSummary()
    Dim wsReport As Worksheet
    Dim totalRevenue As Double
    Dim totalOutstanding As Double
    Dim totalOrders As Long
    Dim completedOrders As Long
    
    Set wsReport = ActiveWorkbook.Worksheets.Add
    wsReport.Name = "Financial_Summary_" & Format(Date, "yyyymmdd")
    
    ' Calculate metrics
    totalRevenue = CalculateTotalRevenue()
    totalOutstanding = CalculateTotalOutstanding()
    totalOrders = CountRows(ActiveWorkbook.Worksheets(WS_ORDERS)) - 1
    completedOrders = CountCompletedOrders()
    
    With wsReport
        .Range("A1").Value = "FINANCIAL SUMMARY REPORT"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "Report Date:"
        .Range("B3").Value = Format(Date, "mmmm dd, yyyy")
        
        .Range("A5").Value = "REVENUE METRICS"
        .Range("A5").Font.Bold = True
        .Range("A6").Value = "Total Revenue:"
        .Range("B6").Value = Format(totalRevenue, "Currency")
        .Range("A7").Value = "Outstanding Payments:"
        .Range("B7").Value = Format(totalOutstanding, "Currency")
        .Range("A8").Value = "Net Revenue:"
        .Range("B8").Value = Format(totalRevenue - totalOutstanding, "Currency")
        
        .Range("A10").Value = "ORDER METRICS"
        .Range("A10").Font.Bold = True
        .Range("A11").Value = "Total Orders:"
        .Range("B11").Value = totalOrders
        .Range("A12").Value = "Completed Orders:"
        .Range("B12").Value = completedOrders
        .Range("A13").Value = "Completion Rate:"
        .Range("B13").Value = Format(completedOrders / totalOrders, "Percent")
        
        .Columns.AutoFit
    End With
    
    MsgBox "Financial summary generated successfully!", vbInformation
End Sub

Function CalculateTotalOutstanding() As Double
    Dim ws As Worksheet
    Dim i As Long
    Dim total As Double
    
    Set ws = ActiveWorkbook.Worksheets(WS_PAYMENTS)
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        total = total + ws.Cells(i, 6).Value ' Outstanding Amount column
    Next i
    
    CalculateTotalOutstanding = total
End Function

' =====================================================
' NAVIGATION MODULE
' =====================================================

Sub CreateNavigationMenu()
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        Call AddNavigationButtons(ws)
    Next ws
End Sub

Sub AddNavigationButtons(ws As Worksheet)
    ' Add navigation buttons to each worksheet
    ' This would typically involve creating actual button objects
    ' For simplicity, we'll add text-based navigation
    
    With ws
        .Range("A" & .Rows.Count).End(xlUp).Offset(2, 0).Value = "Navigation:"
        .Range("A" & .Rows.Count).End(xlUp).Offset(1, 0).Value = "Dashboard | Customers | Orders | Employees | Payments | Reports"
    End With
End Sub

' =====================================================
' SAMPLE DATA GENERATION
' =====================================================

Sub GenerateSampleData()
    Call GenerateSampleCustomers
    Call GenerateSampleEmployees
    Call GenerateSampleOrders
End Sub

Sub GenerateSampleCustomers()
    Dim ws As Worksheet
    Dim i As Long
    Dim customerNames As Variant
    Dim companyNames As Variant
    
    Set ws = ActiveWorkbook.Worksheets(WS_CUSTOMERS)
    
    customerNames = Array("John Smith", "Sarah Johnson", "Michael Brown", "Emily Davis", "David Wilson")
    companyNames = Array("TechCorp Solutions", "Digital Dynamics", "Innovation Labs", "Future Systems", "Creative Agency")
    
    For i = 0 To 4
        With ws
            .Cells(i + 2, 1) = "CUS-" & Format(Now + i, "yyyymmddhhmmss")
            .Cells(i + 2, 2) = customerNames(i)
            .Cells(i + 2, 3) = "ACC-" & (1000 + i)
            .Cells(i + 2, 4) = companyNames(i)
            .Cells(i + 2, 5) = "+1-555-" & (1000 + i * 111)
            .Cells(i + 2, 6) = LCase(Replace(customerNames(i), " ", ".")) & "@" & LCase(Replace(companyNames(i), " ", "")) & ".com"
            .Cells(i + 2, 7) = Date - (30 - i * 5)
            .Cells(i + 2, 8) = Int(Rnd() * 10) + 1
            .Cells(i + 2, 9) = Int(Rnd() * 50000) + 10000
            .Cells(i + 2, 10) = "Active"
        End With
    Next i
End Sub

Sub GenerateSampleEmployees()
    Dim ws As Worksheet
    Dim i As Long
    Dim employeeNames As Variant
    Dim roles As Variant
    Dim responsibilities As Variant
    
    Set ws = ActiveWorkbook.Worksheets(WS_EMPLOYEES)
    
    employeeNames = Array("Alice Chen", "Bob Martinez", "Carol Thompson", "Daniel Kim", "Eva Rodriguez")
    roles = Array("Senior Editor", "Production Manager", "Graphic Designer", "Project Coordinator", "Marketing Specialist")
    responsibilities = Array("Content Editing & Review", "Print Production & Quality Control", "Layout Design & Visual Content", "Project Timeline & Client Communication", "Digital Marketing & Promotion")
    
    For i = 0 To 4
        With ws
            .Cells(i + 2, 1) = "EMP-" & Format(Now + i, "yyyymmddhhmmss")
            .Cells(i + 2, 2) = employeeNames(i)
            .Cells(i + 2, 3) = roles(i)
            .Cells(i + 2, 4) = responsibilities(i)
            .Cells(i + 2, 5) = LCase(Replace(employeeNames(i), " ", ".")) & "@cyberlinkpub.com"
            .Cells(i + 2, 6) = "+1-555-" & (2000 + i * 111)
            .Cells(i + 2, 7) = Date - (365 - i * 30)
            .Cells(i + 2, 8) = Choose(i + 1, "Senior", "Expert", "Intermediate", "Senior", "Intermediate")
            .Cells(i + 2, 9) = Int(Rnd() * 15) + 3
            .Cells(i + 2, 10) = Int(Rnd() * 50) + 20
            .Cells(i + 2, 11) = Round(Rnd() * 2 + 3, 1)
            .Cells(i + 2, 12) = "Active"
        End With
    Next i
End Sub

Sub GenerateSampleOrders()
    Dim ws As Worksheet
    Dim wsCustomers As Worksheet
    Dim wsEmployees As Worksheet
    Dim i As Long
    Dim orderTypes As Variant
    Dim descriptions As Variant
    Dim statuses As Variant
    
    Set ws = ActiveWorkbook.Worksheets(WS_ORDERS)
    Set wsCustomers = ActiveWorkbook.Worksheets(WS_CUSTOMERS)
    Set wsEmployees = ActiveWorkbook.Worksheets(WS_EMPLOYEES)
    
    orderTypes = Array("Book Publishing", "Magazine Layout", "Brochure Design", "Annual Report", "Marketing Materials")
    descriptions = Array("Complete publication with editing and design", "Professional layout with custom graphics", "High-quality print design with branding", "Comprehensive report with data visualization", "Multi-channel marketing campaign materials")
    statuses = Array("Ordered", "In Progress", "On Schedule", "Completed", "Shipped")
    
    For i = 0 To 4
        Dim quantity As Long
        Dim unitPrice As Double
        
        quantity = Int(Rnd() * 1000) + 100
        unitPrice = Round(Rnd() * 50 + 10, 2)
        
        With ws
            .Cells(i + 2, 1) = "ORD-" & Format(Now + i, "yyyymmddhhmmss")
            .Cells(i + 2, 2) = wsCustomers.Cells(i + 2, 1).Value ' Customer ID
            .Cells(i + 2, 3) = wsCustomers.Cells(i + 2, 2).Value ' Customer Name
            .Cells(i + 2, 4) = orderTypes(i)
            .Cells(i + 2, 5) = descriptions(i)
            .Cells(i + 2, 6) = quantity
            .Cells(i + 2, 7) = unitPrice
            .Cells(i + 2, 8) = quantity * unitPrice
            .Cells(i + 2, 9) = Date - (30 - i * 5)
            .Cells(i + 2, 10) = statuses(Int(Rnd() * 5))
            .Cells(i + 2, 11) = Choose(Int(Rnd() * 4) + 1, "Bid", "Advance", "Full Payment", "Credit")
            .Cells(i + 2, 12) = wsEmployees.Cells(i + 2, 2).Value ' Employee Name
            .Cells(i + 2, 13) = Date + (30 + i * 5)
            .Cells(i + 2, 14) = IIf(Rnd() > 0.5, Date + (25 + i * 3), "")
            .Cells(i + 2, 15) = Choose(Int(Rnd() * 4) + 1, "Low", "Medium", "High", "Urgent")
            .Cells(i + 2, 16) = "Sample order notes"
        End With
        
        ' Create corresponding payment record
        Call CreatePaymentRecord(.Cells(i + 2, 1).Value, .Cells(i + 2, 2).Value, quantity * unitPrice)
    Next i
End Sub

' =====================================================
' UTILITY FUNCTIONS
' =====================================================

Sub ShowAbout()
    MsgBox "Cyberlink Publishing Management System" & vbCrLf & _
           "Version: " & VERSION & vbCrLf & _
           "Created with Excel VBA" & vbCrLf & vbCrLf & _
           "Features:" & vbCrLf & _
           "• Customer Management" & vbCrLf & _
           "• Order Tracking" & vbCrLf & _
           "• Employee Management" & vbCrLf & _
           "• Payment Processing" & vbCrLf & _
           "• Comprehensive Reporting" & vbCrLf & _
           "• Data Export Capabilities", vbInformation, "About " & COMPANY_NAME
End Sub

Sub BackupData()
    Dim backupName As String
    backupName = "Cyberlink_Backup_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    
    ActiveWorkbook.SaveCopyAs ThisWorkbook.Path & "\" & backupName
    MsgBox "Data backed up successfully as: " & backupName, vbInformation
End Sub

Sub ClearAllData()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Are you sure you want to clear all data? This action cannot be undone.", vbYesNo + vbCritical)
    
    If response = vbYes Then
        Call SetupCustomersSheet
        Call SetupOrdersSheet
        Call SetupEmployeesSheet
        Call SetupPaymentsSheet
        Call UpdateDashboard
        
        MsgBox "All data cleared successfully!", vbInformation
    End If
End Sub

' =====================================================
' SETTINGS MODULE
' =====================================================

Sub SetupSettingsSheet()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(WS_SETTINGS)
    
    With ws
        .Cells.Clear
        
        .Range("A1").Value = "SYSTEM SETTINGS"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "Company Information:"
        .Range("A4").Value = "Company Name:"
        .Range("B4").Value = COMPANY_NAME
        .Range("A5").Value = "Version:"
        .Range("B5").Value = VERSION
        
        .Range("A7").Value = "System Actions:"
        .Range("A8").Value = "• Initialize System"
        .Range("A9").Value = "• Generate Sample Data"
        .Range("A10").Value = "• Backup Data"
        .Range("A11").Value = "• Clear All Data"
        .Range("A12").Value = "• Show About"
        
        .Range("A14").Value = "Report Generation:"
        .Range("A15").Value = "• Customer Report"
        .Range("A16").Value = "• Orders Report"
        .Range("A17").Value = "• Payments Report"
        .Range("A18").Value = "• Financial Summary"
    End With
End Sub