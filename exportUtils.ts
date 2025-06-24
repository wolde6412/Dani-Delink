import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import 'jspdf-autotable';
import { Customer, Employee, Order, Payment } from '../types';
import { format } from 'date-fns';

declare module 'jspdf' {
  interface jsPDF {
    autoTable: (options: any) => jsPDF;
  }
}

export const exportToExcel = (data: any[], filename: string, sheetName: string) => {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  XLSX.writeFile(workbook, `${filename}.xlsx`);
};

export const exportToPDF = (data: any[], columns: string[], title: string, filename: string) => {
  const doc = new jsPDF();
  
  // Add title
  doc.setFontSize(16);
  doc.text(title, 14, 22);
  
  // Add date
  doc.setFontSize(10);
  doc.text(`Generated on: ${format(new Date(), 'PPP')}`, 14, 30);
  
  // Add table
  doc.autoTable({
    head: [columns],
    body: data.map(item => columns.map(col => {
      if (col.toLowerCase().includes('date') && item[col]) {
        return format(new Date(item[col]), 'PP');
      }
      if (typeof item[col] === 'number' && col.toLowerCase().includes('price')) {
        return `$${item[col].toFixed(2)}`;
      }
      return item[col] || 'N/A';
    })),
    startY: 35,
    styles: { fontSize: 8 },
    headStyles: { fillColor: [37, 99, 235] }
  });
  
  doc.save(`${filename}.pdf`);
};

export const exportToCSV = (data: any[], filename: string) => {
  if (data.length === 0) return;
  
  const headers = Object.keys(data[0]);
  const csvContent = [
    headers.join(','),
    ...data.map(row => 
      headers.map(header => {
        const value = row[header];
        if (typeof value === 'string' && value.includes(',')) {
          return `"${value}"`;
        }
        return value || '';
      }).join(',')
    )
  ].join('\n');
  
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  link.setAttribute('href', url);
  link.setAttribute('download', `${filename}.csv`);
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

export const formatCustomersForExport = (customers: Customer[]) => {
  return customers.map(customer => ({
    'Customer ID': customer.id,
    'Name': customer.name,
    'Account Number': customer.accountNumber,
    'Company Name': customer.companyName,
    'Contact Number': customer.contactNumber,
    'Email': customer.email,
    'Created Date': format(customer.createdAt, 'PP')
  }));
};

export const formatOrdersForExport = (orders: Order[]) => {
  return orders.map(order => ({
    'Order ID': order.id,
    'Customer Name': order.customerName,
    'Order Type': order.orderType,
    'Description': order.description,
    'Quantity': order.quantity,
    'Unit Price': `$${order.unitPrice.toFixed(2)}`,
    'Total Price': `$${order.totalPrice.toFixed(2)}`,
    'Order Date': format(order.orderDate, 'PP'),
    'Status': order.status,
    'Payment Status': order.paymentStatus,
    'Assigned Employee': order.assignedEmployeeName,
    'Estimated Delivery': format(order.estimatedDelivery, 'PP'),
    'Actual Delivery': order.actualDelivery ? format(order.actualDelivery, 'PP') : 'Pending'
  }));
};

export const formatPaymentsForExport = (payments: Payment[]) => {
  return payments.map(payment => ({
    'Payment ID': payment.id,
    'Order ID': payment.orderId,
    'Total Amount': `$${payment.amount.toFixed(2)}`,
    'Paid Amount': `$${payment.paidAmount.toFixed(2)}`,
    'Outstanding Amount': `$${payment.outstandingAmount.toFixed(2)}`,
    'Payment Date': payment.paymentDate ? format(payment.paymentDate, 'PP') : 'Not Paid',
    'Status': payment.status,
    'Payment Method': payment.paymentMethod || 'N/A'
  }));
};

export const formatEmployeesForExport = (employees: Employee[]) => {
  return employees.map(employee => ({
    'Employee ID': employee.id,
    'Name': employee.name,
    'Role': employee.role,
    'Responsibility': employee.responsibility,
    'Assigned Orders': employee.assignedOrders,
    'Completed Tasks': employee.completedTasks,
    'Performance Rating': employee.performanceRating.toFixed(1)
  }));
};