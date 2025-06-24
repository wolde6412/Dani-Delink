export interface Customer {
  id: string;
  name: string;
  accountNumber: string;
  companyName: string;
  contactNumber: string;
  email: string;
  createdAt: Date;
  totalOrders: number;
  totalSpent: number;
  lastOrderDate?: Date;
  status: 'active' | 'inactive';
}

export interface Employee {
  id: string;
  name: string;
  role: string;
  responsibility: string;
  assignedOrders: number;
  completedTasks: number;
  performanceRating: number;
  skillLevel: 'junior' | 'intermediate' | 'senior' | 'expert';
  workload: number;
  isActive: boolean;
  joinDate: Date;
  email: string;
  phone: string;
}

export interface Order {
  id: string;
  customerId: string;
  customerName: string;
  orderType: string;
  description: string;
  quantity: number;
  unitPrice: number;
  totalPrice: number;
  orderDate: Date;
  status: 'ordered' | 'in-progress' | 'on-schedule' | 'completed' | 'shipped';
  paymentStatus: 'bid' | 'advance' | 'full-payment' | 'credit';
  assignedEmployeeId: string;
  assignedEmployeeName: string;
  estimatedDelivery: Date;
  actualDelivery?: Date;
  priority: 'low' | 'medium' | 'high' | 'urgent';
  notes: string;
  confirmations: OrderConfirmation[];
}

export interface Payment {
  id: string;
  customerId: string;
  orderId: string;
  amount: number;
  paidAmount: number;
  outstandingAmount: number;
  paymentDate?: Date;
  status: 'paid-full' | 'paid-advance' | 'unpaid-bid' | 'unpaid-credit';
  paymentMethod?: 'credit-card' | 'bank-transfer' | 'check' | 'cash' | 'online';
  transactionId?: string;
  dueDate?: Date;
  isOverdue: boolean;
  paymentHistory: PaymentTransaction[];
}

export interface PaymentTransaction {
  id: string;
  amount: number;
  date: Date;
  method: string;
  transactionId: string;
  notes?: string;
}

export interface OrderConfirmation {
  id: string;
  type: 'order-received' | 'employee-assigned' | 'in-progress' | 'quality-check' | 'payment-verified' | 'shipped' | 'delivered';
  confirmedBy: string;
  confirmedAt: Date;
  notes?: string;
  digitalSignature?: string;
  notificationSent: boolean;
}

export interface Notification {
  id: string;
  type: 'payment-due' | 'order-overdue' | 'delivery-scheduled' | 'task-assigned' | 'payment-received';
  title: string;
  message: string;
  recipientId: string;
  recipientType: 'customer' | 'employee' | 'admin';
  isRead: boolean;
  createdAt: Date;
  priority: 'low' | 'medium' | 'high';
}

export interface DashboardStats {
  totalRevenue: number;
  totalOrders: number;
  completedOrders: number;
  pendingPayments: number;
  activeEmployees: number;
  averageOrderValue: number;
  overduePayments: number;
  monthlyGrowth: number;
  customerRetentionRate: number;
  onTimeDeliveryRate: number;
}

export interface BusinessMetrics {
  dailyRevenue: { date: string; revenue: number }[];
  ordersByStatus: { status: string; count: number }[];
  paymentsByMethod: { method: string; amount: number }[];
  employeePerformance: { name: string; rating: number; tasks: number }[];
  customerSegmentation: { segment: string; count: number; revenue: number }[];
  monthlyTrends: { month: string; orders: number; revenue: number }[];
}