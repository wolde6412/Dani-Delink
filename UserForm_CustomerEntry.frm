VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_CustomerEntry 
   Caption         =   "Customer Entry Form"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8400
   OleObjectBlob   =   "UserForm_CustomerEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_CustomerEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' =====================================================
' CUSTOMER ENTRY USERFORM
' Professional customer entry interface
' =====================================================

Option Explicit

Private Sub UserForm_Initialize()
    ' Set default values
    txtAccountNumber.Value = "ACC-" & Format(Now, "yyyymmddhhmmss")
    txtCreatedDate.Value = Format(Date, "mm/dd/yyyy")
    cmbStatus.Value = "Active"
    txtTotalOrders.Value = "0"
    txtTotalSpent.Value = "$0.00"
End Sub

Private Sub btnSave_Click()
    ' Validate required fields
    If Not ValidateForm() Then Exit Sub
    
    ' Save customer
    Call SaveCustomer
    
    ' Close form
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function ValidateForm() As Boolean
    ValidateForm = True
    
    If txtName.Value = "" Then
        MsgBox "Please enter customer name.", vbWarning
        txtName.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If txtCompanyName.Value = "" Then
        MsgBox "Please enter company name.", vbWarning
        txtCompanyName.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If txtContactNumber.Value = "" Then
        MsgBox "Please enter contact number.", vbWarning
        txtContactNumber.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If txtEmail.Value = "" Then
        MsgBox "Please enter email address.", vbWarning
        txtEmail.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    ' Validate email format
    If InStr(txtEmail.Value, "@") = 0 Or InStr(txtEmail.Value, ".") = 0 Then
        MsgBox "Please enter a valid email address.", vbWarning
        txtEmail.SetFocus
        ValidateForm = False
        Exit Function
    End If
End Function

Private Sub SaveCustomer()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim customerID As String
    
    Set ws = ActiveWorkbook.Worksheets("Customers")
    
    ' Generate customer ID
    customerID = "CUS-" & Format(Now, "yyyymmddhhmmss")
    
    ' Find next row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Save customer data
    With ws
        .Cells(lastRow, 1) = customerID
        .Cells(lastRow, 2) = txtName.Value
        .Cells(lastRow, 3) = txtAccountNumber.Value
        .Cells(lastRow, 4) = txtCompanyName.Value
        .Cells(lastRow, 5) = txtContactNumber.Value
        .Cells(lastRow, 6) = txtEmail.Value
        .Cells(lastRow, 7) = CDate(txtCreatedDate.Value)
        .Cells(lastRow, 8) = Val(txtTotalOrders.Value)
        .Cells(lastRow, 9) = Val(Replace(txtTotalSpent.Value, "$", ""))
        .Cells(lastRow, 10) = cmbStatus.Value
    End With
    
    ' Update dashboard
    Call UpdateDashboard
    
    MsgBox "Customer saved successfully!" & vbCrLf & "Customer ID: " & customerID, vbInformation
End Sub