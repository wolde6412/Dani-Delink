VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_OrderEntry 
   Caption         =   "Order Entry Form"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12015
   OleObjectBlob   =   "UserForm_OrderEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_OrderEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' =====================================================
' ORDER ENTRY USERFORM
' Professional order entry interface
' =====================================================

Option Explicit

Private Sub UserForm_Initialize()
    ' Initialize form controls
    Call LoadCustomers
    Call LoadEmployees
    Call LoadOrderTypes
    Call SetDefaultValues
End Sub

Private Sub LoadCustomers()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    
    Set ws = ActiveWorkbook.Worksheets("Customers")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    cmbCustomer.Clear
    For i = 2 To lastRow
        cmbCustomer.AddItem ws.Cells(i, 1).Value & " - " & ws.Cells(i, 2).Value
    Next i
End Sub

Private Sub LoadEmployees()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    
    Set ws = ActiveWorkbook.Worksheets("Employees")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    cmbEmployee.Clear
    For i = 2 To lastRow
        cmbEmployee.AddItem ws.Cells(i, 1).Value & " - " & ws.Cells(i, 2).Value
    Next i
End Sub

Private Sub LoadOrderTypes()
    cmbOrderType.Clear
    cmbOrderType.AddItem "Book Publishing"
    cmbOrderType.AddItem "Magazine Layout"
    cmbOrderType.AddItem "Brochure Design"
    cmbOrderType.AddItem "Annual Report"
    cmbOrderType.AddItem "Marketing Materials"
    cmbOrderType.AddItem "Corporate Newsletter"
    cmbOrderType.AddItem "Product Catalog"
    cmbOrderType.AddItem "Training Manual"
End Sub

Private Sub SetDefaultValues()
    txtOrderDate.Value = Format(Date, "mm/dd/yyyy")
    txtEstimatedDelivery.Value = Format(Date + 30, "mm/dd/yyyy")
    txtQuantity.Value = "1"
    txtUnitPrice.Value = "100.00"
    cmbStatus.Value = "Ordered"
    cmbPaymentStatus.Value = "Bid"
    cmbPriority.Value = "Medium"
    
    Call CalculateTotal
End Sub

Private Sub txtQuantity_Change()
    Call CalculateTotal
End Sub

Private Sub txtUnitPrice_Change()
    Call CalculateTotal
End Sub

Private Sub CalculateTotal()
    Dim quantity As Long
    Dim unitPrice As Double
    Dim total As Double
    
    quantity = Val(txtQuantity.Value)
    unitPrice = Val(txtUnitPrice.Value)
    total = quantity * unitPrice
    
    txtTotalPrice.Value = Format(total, "Currency")
End Sub

Private Sub btnSave_Click()
    ' Validate required fields
    If Not ValidateForm() Then Exit Sub
    
    ' Save order
    Call SaveOrder
    
    ' Close form
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function ValidateForm() As Boolean
    ValidateForm = True
    
    If cmbCustomer.Value = "" Then
        MsgBox "Please select a customer.", vbWarning
        cmbCustomer.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If cmbOrderType.Value = "" Then
        MsgBox "Please select an order type.", vbWarning
        cmbOrderType.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If txtDescription.Value = "" Then
        MsgBox "Please enter a description.", vbWarning
        txtDescription.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If Val(txtQuantity.Value) <= 0 Then
        MsgBox "Please enter a valid quantity.", vbWarning
        txtQuantity.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If Val(txtUnitPrice.Value) <= 0 Then
        MsgBox "Please enter a valid unit price.", vbWarning
        txtUnitPrice.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If cmbEmployee.Value = "" Then
        MsgBox "Please select an employee.", vbWarning
        cmbEmployee.SetFocus
        ValidateForm = False
        Exit Function
    End If
End Function

Private Sub SaveOrder()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim orderID As String
    Dim customerID As String
    Dim employeeID As String
    
    Set ws = ActiveWorkbook.Worksheets("Orders")
    
    ' Generate order ID
    orderID = "ORD-" & Format(Now, "yyyymmddhhmmss")
    
    ' Extract IDs from combo boxes
    customerID = Split(cmbCustomer.Value, " - ")(0)
    employeeID = Split(cmbEmployee.Value, " - ")(0)
    
    ' Find next row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Save order data
    With ws
        .Cells(lastRow, 1) = orderID
        .Cells(lastRow, 2) = customerID
        .Cells(lastRow, 3) = Split(cmbCustomer.Value, " - ")(1)
        .Cells(lastRow, 4) = cmbOrderType.Value
        .Cells(lastRow, 5) = txtDescription.Value
        .Cells(lastRow, 6) = Val(txtQuantity.Value)
        .Cells(lastRow, 7) = Val(txtUnitPrice.Value)
        .Cells(lastRow, 8) = Val(txtQuantity.Value) * Val(txtUnitPrice.Value)
        .Cells(lastRow, 9) = CDate(txtOrderDate.Value)
        .Cells(lastRow, 10) = cmbStatus.Value
        .Cells(lastRow, 11) = cmbPaymentStatus.Value
        .Cells(lastRow, 12) = Split(cmbEmployee.Value, " - ")(1)
        .Cells(lastRow, 13) = CDate(txtEstimatedDelivery.Value)
        .Cells(lastRow, 14) = IIf(txtActualDelivery.Value <> "", CDate(txtActualDelivery.Value), "")
        .Cells(lastRow, 15) = cmbPriority.Value
        .Cells(lastRow, 16) = txtNotes.Value
    End With
    
    ' Create payment record
    Call CreatePaymentRecord(orderID, customerID, Val(txtQuantity.Value) * Val(txtUnitPrice.Value))
    
    ' Update dashboard
    Call UpdateDashboard
    
    MsgBox "Order saved successfully!" & vbCrLf & "Order ID: " & orderID, vbInformation
End Sub