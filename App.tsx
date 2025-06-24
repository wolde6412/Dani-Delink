import React, { useState, useEffect } from 'react';
import Sidebar from './components/Layout/Sidebar';
import Header from './components/Layout/Header';
import Dashboard from './components/Dashboard/Dashboard';
import CustomerList from './components/Customers/CustomerList';
import OrderList from './components/Orders/OrderList';
import Reports from './components/Reports/Reports';
import { Customer, Employee, Order, Payment } from './types';
import { generateMockData, generateId } from './utils/dataGenerator';

function App() {
  const [activeTab, setActiveTab] = useState('dashboard');
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [orders, setOrders] = useState<Order[]>([]);
  const [payments, setPayments] = useState<Payment[]>([]);

  useEffect(() => {
    const mockData = generateMockData();
    setCustomers(mockData.customers);
    setEmployees(mockData.employees);
    setOrders(mockData.orders);
    setPayments(mockData.payments);
  }, []);

  const handleAddCustomer = (customerData: Omit<Customer, 'id' | 'createdAt'>) => {
    const newCustomer: Customer = {
      ...customerData,
      id: generateId('CUS'),
      createdAt: new Date()
    };
    setCustomers(prev => [newCustomer, ...prev]);
  };

  const handleUpdateCustomer = (id: string, customerData: Partial<Customer>) => {
    setCustomers(prev => prev.map(customer => 
      customer.id === id ? { ...customer, ...customerData } : customer
    ));
  };

  const handleDeleteCustomer = (id: string) => {
    setCustomers(prev => prev.filter(customer => customer.id !== id));
  };

  const handleAddOrder = (orderData: Omit<Order, 'id' | 'orderDate'>) => {
    const newOrder: Order = {
      ...orderData,
      id: generateId('ORD'),
      orderDate: new Date()
    };
    setOrders(prev => [newOrder, ...prev]);
    
    // Create corresponding payment record
    const newPayment: Payment = {
      id: generateId('PAY'),
      customerId: orderData.customerId,
      orderId: newOrder.id,
      amount: orderData.totalPrice,
      paidAmount: orderData.paymentStatus === 'full-payment' ? orderData.totalPrice : 0,
      outstandingAmount: orderData.paymentStatus === 'full-payment' ? 0 : orderData.totalPrice,
      paymentDate: orderData.paymentStatus === 'full-payment' ? new Date() : undefined,
      status: orderData.paymentStatus === 'full-payment' ? 'paid-full' :
              orderData.paymentStatus === 'advance' ? 'paid-advance' :
              orderData.paymentStatus === 'bid' ? 'unpaid-bid' : 'unpaid-credit'
    };
    setPayments(prev => [newPayment, ...prev]);
  };

  const handleUpdateOrder = (id: string, orderData: Partial<Order>) => {
    setOrders(prev => prev.map(order => 
      order.id === id ? { ...order, ...orderData } : order
    ));
    
    // Update corresponding payment if payment status changed
    if (orderData.paymentStatus) {
      setPayments(prev => prev.map(payment => {
        if (payment.orderId === id) {
          let paidAmount = payment.paidAmount;
          let status = payment.status;
          
          if (orderData.paymentStatus === 'full-payment') {
            paidAmount = payment.amount;
            status = 'paid-full';
          } else if (orderData.paymentStatus === 'advance') {
            status = 'paid-advance';
          } else if (orderData.paymentStatus === 'bid') {
            status = 'unpaid-bid';
            paidAmount = 0;
          } else if (orderData.paymentStatus === 'credit') {
            status = 'unpaid-credit';
            paidAmount = 0;
          }
          
          return {
            ...payment,
            paidAmount,
            outstandingAmount: payment.amount - paidAmount,
            status,
            paymentDate: paidAmount > 0 ? (payment.paymentDate || new Date()) : undefined
          };
        }
        return payment;
      }));
    }
  };

  const renderContent = () => {
    switch (activeTab) {
      case 'dashboard':
        return (
          <Dashboard
            customers={customers}
            employees={employees}
            orders={orders}
            payments={payments}
          />
        );
      case 'customers':
        return (
          <CustomerList
            customers={customers}
            onAddCustomer={handleAddCustomer}
            onUpdateCustomer={handleUpdateCustomer}
            onDeleteCustomer={handleDeleteCustomer}
          />
        );
      case 'orders':
        return (
          <OrderList
            orders={orders}
            customers={customers}
            employees={employees}
            onAddOrder={handleAddOrder}
            onUpdateOrder={handleUpdateOrder}
          />
        );
      case 'employees':
        return (
          <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
            <h2 className="text-xl font-semibold text-gray-900 mb-4">Employee Management</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {employees.map(employee => (
                <div key={employee.id} className="p-4 border border-gray-200 rounded-lg">
                  <h3 className="font-semibold text-gray-900">{employee.name}</h3>
                  <p className="text-sm text-gray-600">{employee.role}</p>
                  <p className="text-xs text-gray-500 mt-1">{employee.responsibility}</p>
                  <div className="mt-3 flex justify-between text-sm">
                    <span>Orders: {employee.assignedOrders}</span>
                    <span>Rating: {employee.performanceRating.toFixed(1)}/5</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        );
      case 'payments':
        return (
          <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
            <h2 className="text-xl font-semibold text-gray-900 mb-4">Payment Tracking</h2>
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead className="bg-gray-50">
                  <tr>
                    <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase">Payment ID</th>
                    <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase">Order ID</th>
                    <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase">Amount</th>
                    <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase">Paid</th>
                    <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase">Outstanding</th>
                    <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-200">
                  {payments.slice(0, 20).map(payment => (
                    <tr key={payment.id} className="hover:bg-gray-50">
                      <td className="px-4 py-3 text-sm font-mono">{payment.id}</td>
                      <td className="px-4 py-3 text-sm font-mono">{payment.orderId}</td>
                      <td className="px-4 py-3 text-sm font-medium">${payment.amount.toFixed(2)}</td>
                      <td className="px-4 py-3 text-sm text-green-600 font-medium">${payment.paidAmount.toFixed(2)}</td>
                      <td className="px-4 py-3 text-sm text-red-600 font-medium">${payment.outstandingAmount.toFixed(2)}</td>
                      <td className="px-4 py-3">
                        <span className={`px-2 py-1 text-xs font-semibold rounded-full ${
                          payment.status === 'paid-full' ? 'bg-green-100 text-green-800' :
                          payment.status === 'paid-advance' ? 'bg-blue-100 text-blue-800' :
                          'bg-red-100 text-red-800'
                        }`}>
                          {payment.status.replace('-', ' ')}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        );
      case 'reports':
        return (
          <Reports
            customers={customers}
            employees={employees}
            orders={orders}
            payments={payments}
          />
        );
      default:
        return null;
    }
  };

  return (
    <div className="flex h-screen bg-gray-50">
      <Sidebar activeTab={activeTab} setActiveTab={setActiveTab} />
      <div className="flex-1 flex flex-col overflow-hidden">
        <Header activeTab={activeTab} />
        <main className="flex-1 overflow-auto p-6">
          {renderContent()}
        </main>
      </div>
    </div>
  );
}

export default App;