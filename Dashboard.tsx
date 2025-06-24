import React from 'react';
import { DollarSign, ShoppingCart, Users, CreditCard, TrendingUp, Package } from 'lucide-react';
import StatsCard from './StatsCard';
import ChartCard from './ChartCard';
import { Customer, Employee, Order, Payment, DashboardStats } from '../../types';

interface DashboardProps {
  customers: Customer[];
  employees: Employee[];
  orders: Order[];
  payments: Payment[];
}

const Dashboard: React.FC<DashboardProps> = ({ customers, employees, orders, payments }) => {
  // Calculate dashboard stats
  const stats: DashboardStats = {
    totalRevenue: payments.reduce((sum, p) => sum + p.paidAmount, 0),
    totalOrders: orders.length,
    completedOrders: orders.filter(o => o.status === 'completed' || o.status === 'shipped').length,
    pendingPayments: payments.filter(p => p.outstandingAmount > 0).length,
    activeEmployees: employees.length,
    averageOrderValue: orders.reduce((sum, o) => sum + o.totalPrice, 0) / orders.length
  };

  // Chart data
  const orderStatusData = [
    { name: 'Ordered', value: orders.filter(o => o.status === 'ordered').length },
    { name: 'In Progress', value: orders.filter(o => o.status === 'in-progress').length },
    { name: 'On Schedule', value: orders.filter(o => o.status === 'on-schedule').length },
    { name: 'Completed', value: orders.filter(o => o.status === 'completed').length },
    { name: 'Shipped', value: orders.filter(o => o.status === 'shipped').length }
  ];

  const paymentStatusData = [
    { name: 'Paid Full', value: payments.filter(p => p.status === 'paid-full').length },
    { name: 'Paid Advance', value: payments.filter(p => p.status === 'paid-advance').length },
    { name: 'Unpaid Bid', value: payments.filter(p => p.status === 'unpaid-bid').length },
    { name: 'Unpaid Credit', value: payments.filter(p => p.status === 'unpaid-credit').length }
  ];

  const revenueData = Array.from({ length: 7 }, (_, i) => {
    const date = new Date();
    date.setDate(date.getDate() - (6 - i));
    const dayRevenue = payments
      .filter(p => p.paymentDate && new Date(p.paymentDate).toDateString() === date.toDateString())
      .reduce((sum, p) => sum + p.paidAmount, 0);
    return {
      name: date.toLocaleDateString('en-US', { weekday: 'short' }),
      value: dayRevenue
    };
  });

  const topEmployeesData = employees
    .sort((a, b) => b.completedTasks - a.completedTasks)
    .slice(0, 5)
    .map(emp => ({
      name: emp.name.split(' ')[0],
      value: emp.completedTasks
    }));

  return (
    <div className="space-y-6">
      {/* Stats Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-6">
        <StatsCard
          title="Total Revenue"
          value={`$${stats.totalRevenue.toLocaleString()}`}
          change="+12.5% from last month"
          changeType="positive"
          icon={DollarSign}
          color="bg-blue-600"
          description="Total payments received"
        />
        <StatsCard
          title="Total Orders"
          value={stats.totalOrders}
          change="+8.2% from last month"
          changeType="positive"
          icon={ShoppingCart}
          color="bg-teal-600"
          description="All orders placed"
        />
        <StatsCard
          title="Completed Orders"
          value={stats.completedOrders}
          change="+15.3% from last month"
          changeType="positive"
          icon={Package}
          color="bg-green-600"
          description="Successfully delivered"
        />
        <StatsCard
          title="Pending Payments"
          value={stats.pendingPayments}
          change="-5.1% from last month"
          changeType="positive"
          icon={CreditCard}
          color="bg-orange-600"
          description="Outstanding payments"
        />
        <StatsCard
          title="Active Employees"
          value={stats.activeEmployees}
          icon={Users}
          color="bg-purple-600"
          description="Working on projects"
        />
        <StatsCard
          title="Avg Order Value"
          value={`$${stats.averageOrderValue.toFixed(0)}`}
          change="+3.2% from last month"
          changeType="positive"
          icon={TrendingUp}
          color="bg-indigo-600"
          description="Per order revenue"
        />
      </div>

      {/* Charts */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <ChartCard
          title="Order Status Distribution"
          type="pie"
          data={orderStatusData}
          dataKey="value"
          nameKey="name"
        />
        <ChartCard
          title="Payment Status Overview"
          type="bar"
          data={paymentStatusData}
          dataKey="value"
          nameKey="name"
        />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <ChartCard
          title="Weekly Revenue Trend"
          type="line"
          data={revenueData}
          dataKey="value"
          nameKey="name"
        />
        <ChartCard
          title="Top Performing Employees"
          type="bar"
          data={topEmployeesData}
          dataKey="value"
          nameKey="name"
          colors={['#7C3AED']}
        />
      </div>

      {/* Recent Activity */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
        <h3 className="text-lg font-semibold text-gray-900 mb-4">Recent Orders</h3>
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Order ID</th>
                <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Customer</th>
                <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Type</th>
                <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Amount</th>
                <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status</th>
                <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Employee</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-200">
              {orders.slice(0, 8).map((order) => (
                <tr key={order.id} className="hover:bg-gray-50">
                  <td className="px-4 py-3 text-sm font-medium text-gray-900">{order.id}</td>
                  <td className="px-4 py-3 text-sm text-gray-900">{order.customerName}</td>
                  <td className="px-4 py-3 text-sm text-gray-600">{order.orderType}</td>
                  <td className="px-4 py-3 text-sm font-medium text-gray-900">${order.totalPrice.toFixed(2)}</td>
                  <td className="px-4 py-3">
                    <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
                      order.status === 'completed' || order.status === 'shipped' 
                        ? 'bg-green-100 text-green-800'
                        : order.status === 'in-progress' || order.status === 'on-schedule'
                        ? 'bg-blue-100 text-blue-800'
                        : 'bg-yellow-100 text-yellow-800'
                    }`}>
                      {order.status.replace('-', ' ')}
                    </span>
                  </td>
                  <td className="px-4 py-3 text-sm text-gray-600">{order.assignedEmployeeName}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;