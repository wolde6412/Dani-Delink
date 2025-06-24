import React, { useState } from 'react';
import { Download, FileSpreadsheet, FileText, FileBarChart } from 'lucide-react';
import { Customer, Employee, Order, Payment } from '../../types';
import Button from '../Common/Button';
import ChartCard from '../Dashboard/ChartCard';
import {
  exportToExcel,
  exportToPDF,
  exportToCSV,
  formatCustomersForExport,
  formatOrdersForExport,
  formatPaymentsForExport,
  formatEmployeesForExport
} from '../../utils/exportUtils';

interface ReportsProps {
  customers: Customer[];
  employees: Employee[];
  orders: Order[];
  payments: Payment[];
}

type ExportFormat = 'excel' | 'pdf' | 'csv';
type ReportType = 'customers' | 'orders' | 'payments' | 'employees' | 'summary';

const Reports: React.FC<ReportsProps> = ({ customers, employees, orders, payments }) => {
  const [exportFormat, setExportFormat] = useState<ExportFormat>('excel');
  const [reportType, setReportType] = useState<ReportType>('summary');

  const handleExport = () => {
    const timestamp = new Date().toISOString().split('T')[0];
    
    switch (reportType) {
      case 'customers':
        const customerData = formatCustomersForExport(customers);
        if (exportFormat === 'excel') {
          exportToExcel(customerData, `customers-report-${timestamp}`, 'Customers');
        } else if (exportFormat === 'pdf') {
          exportToPDF(customerData, Object.keys(customerData[0] || {}), 'Customer Report', `customers-report-${timestamp}`);
        } else {
          exportToCSV(customerData, `customers-report-${timestamp}`);
        }
        break;
        
      case 'orders':
        const orderData = formatOrdersForExport(orders);
        if (exportFormat === 'excel') {
          exportToExcel(orderData, `orders-report-${timestamp}`, 'Orders');
        } else if (exportFormat === 'pdf') {
          exportToPDF(orderData, Object.keys(orderData[0] || {}), 'Orders Report', `orders-report-${timestamp}`);
        } else {
          exportToCSV(orderData, `orders-report-${timestamp}`);
        }
        break;
        
      case 'payments':
        const paymentData = formatPaymentsForExport(payments);
        if (exportFormat === 'excel') {
          exportToExcel(paymentData, `payments-report-${timestamp}`, 'Payments');
        } else if (exportFormat === 'pdf') {
          exportToPDF(paymentData, Object.keys(paymentData[0] || {}), 'Payments Report', `payments-report-${timestamp}`);
        } else {
          exportToCSV(paymentData, `payments-report-${timestamp}`);
        }
        break;
        
      case 'employees':
        const employeeData = formatEmployeesForExport(employees);
        if (exportFormat === 'excel') {
          exportToExcel(employeeData, `employees-report-${timestamp}`, 'Employees');
        } else if (exportFormat === 'pdf') {
          exportToPDF(employeeData, Object.keys(employeeData[0] || {}), 'Employee Report', `employees-report-${timestamp}`);
        } else {
          exportToCSV(employeeData, `employees-report-${timestamp}`);
        }
        break;
        
      case 'summary':
        const summaryData = [
          {
            'Metric': 'Total Customers',
            'Value': customers.length,
            'Details': 'Active customer accounts'
          },
          {
            'Metric': 'Total Orders',
            'Value': orders.length,
            'Details': 'All orders placed'
          },
          {
            'Metric': 'Total Revenue',
            'Value': `$${payments.reduce((sum, p) => sum + p.paidAmount, 0).toFixed(2)}`,
            'Details': 'Total payments received'
          },
          {
            'Metric': 'Pending Payments',
            'Value': payments.filter(p => p.outstandingAmount > 0).length,
            'Details': 'Orders with outstanding balance'
          },
          {
            'Metric': 'Active Employees',
            'Value': employees.length,
            'Details': 'Current workforce'
          },
          {
            'Metric': 'Completed Orders',
            'Value': orders.filter(o => o.status === 'completed' || o.status === 'shipped').length,
            'Details': 'Successfully delivered orders'
          }
        ];
        
        if (exportFormat === 'excel') {
          exportToExcel(summaryData, `business-summary-${timestamp}`, 'Summary');
        } else if (exportFormat === 'pdf') {
          exportToPDF(summaryData, ['Metric', 'Value', 'Details'], 'Business Summary Report', `business-summary-${timestamp}`);
        } else {
          exportToCSV(summaryData, `business-summary-${timestamp}`);
        }
        break;
    }
  };

  // Analytics data
  const monthlyRevenue = Array.from({ length: 12 }, (_, i) => {
    const month = new Date();
    month.setMonth(i);
    const monthRevenue = payments
      .filter(p => p.paymentDate && new Date(p.paymentDate).getMonth() === i)
      .reduce((sum, p) => sum + p.paidAmount, 0);
    return {
      name: month.toLocaleDateString('en-US', { month: 'short' }),
      value: monthRevenue
    };
  });

  const orderTypeDistribution = [
    { name: 'Book Publishing', value: orders.filter(o => o.orderType === 'Book Publishing').length },
    { name: 'Magazine Layout', value: orders.filter(o => o.orderType === 'Magazine Layout').length },
    { name: 'Brochure Design', value: orders.filter(o => o.orderType === 'Brochure Design').length },
    { name: 'Annual Report', value: orders.filter(o => o.orderType === 'Annual Report').length },
    { name: 'Marketing Materials', value: orders.filter(o => o.orderType === 'Marketing Materials').length },
    { name: 'Other', value: orders.filter(o => !['Book Publishing', 'Magazine Layout', 'Brochure Design', 'Annual Report', 'Marketing Materials'].includes(o.orderType)).length }
  ].filter(item => item.value > 0);

  const employeePerformance = employees.map(emp => ({
    name: emp.name.split(' ')[0],
    value: emp.performanceRating
  }));

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <div>
          <h2 className="text-xl font-semibold text-gray-900">Reports & Analytics</h2>
          <p className="text-gray-600">Generate comprehensive reports and export data</p>
        </div>
      </div>

      {/* Export Controls */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
        <h3 className="text-lg font-semibold text-gray-900 mb-4">Export Data</h3>
        
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Report Type</label>
            <select
              value={reportType}
              onChange={(e) => setReportType(e.target.value as ReportType)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
            >
              <option value="summary">Business Summary</option>
              <option value="customers">Customer Report</option>
              <option value="orders">Orders Report</option>
              <option value="payments">Payments Report</option>
              <option value="employees">Employee Report</option>
            </select>
          </div>
          
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Export Format</label>
            <select
              value={exportFormat}
              onChange={(e) => setExportFormat(e.target.value as ExportFormat)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
            >
              <option value="excel">Excel (.xlsx)</option>
              <option value="pdf">PDF (.pdf)</option>
              <option value="csv">CSV (.csv)</option>
            </select>
          </div>
          
          <div className="flex items-end">
            <Button
              onClick={handleExport}
              icon={Download}
              className="w-full"
            >
              Export Report
            </Button>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div className="flex items-center p-4 bg-blue-50 rounded-lg">
            <FileSpreadsheet className="w-8 h-8 text-blue-600 mr-3" />
            <div>
              <h4 className="font-medium text-blue-900">Excel Reports</h4>
              <p className="text-sm text-blue-700">Structured data with formulas</p>
            </div>
          </div>
          
          <div className="flex items-center p-4 bg-red-50 rounded-lg">
            <FileText className="w-8 h-8 text-red-600 mr-3" />
            <div>
              <h4 className="font-medium text-red-900">PDF Reports</h4>
              <p className="text-sm text-red-700">Professional formatted documents</p>
            </div>
          </div>
          
          <div className="flex items-center p-4 bg-green-50 rounded-lg">
            <FileBarChart className="w-8 h-8 text-green-600 mr-3" />
            <div>
              <h4 className="font-medium text-green-900">CSV Reports</h4>
              <p className="text-sm text-green-700">Raw data for analysis</p>
            </div>
          </div>
        </div>
      </div>

      {/* Analytics Charts */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <ChartCard
          title="Monthly Revenue Trend"
          type="line"
          data={monthlyRevenue}
          dataKey="value"
          nameKey="name"
        />
        <ChartCard
          title="Order Type Distribution"
          type="pie"
          data={orderTypeDistribution}
          dataKey="value"
          nameKey="name"
        />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <ChartCard
          title="Employee Performance Rating"
          type="bar"
          data={employeePerformance}
          dataKey="value"
          nameKey="name"
          colors={['#10B981']}
        />
        <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">Quick Stats</h3>
          <div className="space-y-4">
            <div className="flex justify-between items-center">
              <span className="text-gray-600">Total Revenue</span>
              <span className="font-semibold">${payments.reduce((sum, p) => sum + p.paidAmount, 0).toFixed(2)}</span>
            </div>
            <div className="flex justify-between items-center">
              <span className="text-gray-600">Average Order Value</span>
              <span className="font-semibold">${(orders.reduce((sum, o) => sum + o.totalPrice, 0) / orders.length).toFixed(2)}</span>
            </div>
            <div className="flex justify-between items-center">
              <span className="text-gray-600">Orders This Month</span>
              <span className="font-semibold">{orders.filter(o => new Date(o.orderDate).getMonth() === new Date().getMonth()).length}</span>
            </div>
            <div className="flex justify-between items-center">
              <span className="text-gray-600">Customer Retention</span>
              <span className="font-semibold text-green-600">94.2%</span>
            </div>
            <div className="flex justify-between items-center">
              <span className="text-gray-600">On-Time Delivery</span>
              <span className="font-semibold text-green-600">96.8%</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default Reports;