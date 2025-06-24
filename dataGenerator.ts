import { Customer, Employee, Order, Payment } from '../types';

const generateId = (prefix: string): string => {
  const timestamp = Date.now().toString(36);
  const random = Math.random().toString(36).substr(2, 5);
  return `${prefix}-${timestamp}-${random}`.toUpperCase();
};

const generateCustomers = (): Customer[] => {
  const companies = [
    'TechCorp Solutions', 'Digital Dynamics', 'Innovation Labs', 'Future Systems',
    'Creative Agency', 'Business Solutions Inc', 'Global Enterprises', 'Smart Industries'
  ];
  
  const names = [
    'John Smith', 'Sarah Johnson', 'Michael Brown', 'Emily Davis',
    'David Wilson', 'Jessica Martinez', 'Christopher Garcia', 'Amanda Rodriguez'
  ];

  return Array.from({ length: 25 }, (_, i) => ({
    id: generateId('CUS'),
    name: names[i % names.length],
    accountNumber: `ACC-${(1000 + i).toString()}`,
    companyName: companies[i % companies.length],
    contactNumber: `+1-555-${(Math.random() * 9000 + 1000).toFixed(0)}`,
    email: `${names[i % names.length].toLowerCase().replace(' ', '.')}@${companies[i % companies.length].toLowerCase().replace(/\s+/g, '')}.com`,
    createdAt: new Date(Date.now() - Math.random() * 90 * 24 * 60 * 60 * 1000)
  }));
};

const generateEmployees = (): Employee[] => {
  const employees = [
    { name: 'Alice Chen', role: 'Senior Editor', responsibility: 'Content Editing & Review' },
    { name: 'Bob Martinez', role: 'Production Manager', responsibility: 'Print Production & Quality Control' },
    { name: 'Carol Thompson', role: 'Graphic Designer', responsibility: 'Layout Design & Visual Content' },
    { name: 'Daniel Kim', role: 'Project Coordinator', responsibility: 'Project Timeline & Client Communication' },
    { name: 'Eva Rodriguez', role: 'Marketing Specialist', responsibility: 'Digital Marketing & Promotion' },
    { name: 'Frank Johnson', role: 'Sales Representative', responsibility: 'Client Acquisition & Account Management' },
    { name: 'Grace Liu', role: 'Quality Assurance', responsibility: 'Final Review & Error Checking' },
    { name: 'Henry Williams', role: 'Customer Service', responsibility: 'Client Support & Order Management' }
  ];

  return employees.map(emp => ({
    id: generateId('EMP'),
    name: emp.name,
    role: emp.role,
    responsibility: emp.responsibility,
    assignedOrders: Math.floor(Math.random() * 15) + 3,
    completedTasks: Math.floor(Math.random() * 50) + 20,
    performanceRating: Math.random() * 2 + 3 // 3-5 rating
  }));
};

const generateOrders = (customers: Customer[], employees: Employee[]): Order[] => {
  const orderTypes = [
    'Book Publishing', 'Magazine Layout', 'Brochure Design', 'Annual Report',
    'Marketing Materials', 'Corporate Newsletter', 'Product Catalog', 'Training Manual'
  ];

  const descriptions = [
    'Complete publication with editing and design',
    'Professional layout with custom graphics',
    'High-quality print design with branding',
    'Comprehensive report with data visualization',
    'Multi-channel marketing campaign materials',
    'Monthly newsletter with custom content',
    'Product showcase with photography',
    'Training documentation with illustrations'
  ];

  const statuses: Order['status'][] = ['ordered', 'in-progress', 'on-schedule', 'completed', 'shipped'];
  const paymentStatuses: Order['paymentStatus'][] = ['bid', 'advance', 'full-payment', 'credit'];

  return Array.from({ length: 50 }, (_, i) => {
    const customer = customers[i % customers.length];
    const employee = employees[i % employees.length];
    const quantity = Math.floor(Math.random() * 1000) + 100;
    const unitPrice = Math.random() * 50 + 10;
    const orderDate = new Date(Date.now() - Math.random() * 60 * 24 * 60 * 60 * 1000);
    
    return {
      id: generateId('ORD'),
      customerId: customer.id,
      customerName: customer.name,
      orderType: orderTypes[i % orderTypes.length],
      description: descriptions[i % descriptions.length],
      quantity,
      unitPrice: Math.round(unitPrice * 100) / 100,
      totalPrice: Math.round(quantity * unitPrice * 100) / 100,
      orderDate,
      status: statuses[Math.floor(Math.random() * statuses.length)],
      paymentStatus: paymentStatuses[Math.floor(Math.random() * paymentStatuses.length)],
      assignedEmployeeId: employee.id,
      assignedEmployeeName: employee.name,
      estimatedDelivery: new Date(orderDate.getTime() + Math.random() * 30 * 24 * 60 * 60 * 1000),
      actualDelivery: Math.random() > 0.5 ? new Date(orderDate.getTime() + Math.random() * 25 * 24 * 60 * 60 * 1000) : undefined
    };
  });
};

const generatePayments = (orders: Order[]): Payment[] => {
  return orders.map(order => {
    const paidPercentage = order.paymentStatus === 'full-payment' ? 1 : 
                          order.paymentStatus === 'advance' ? Math.random() * 0.5 + 0.3 : 0;
    const paidAmount = Math.round(order.totalPrice * paidPercentage * 100) / 100;
    
    return {
      id: generateId('PAY'),
      customerId: order.customerId,
      orderId: order.id,
      amount: order.totalPrice,
      paidAmount,
      outstandingAmount: Math.round((order.totalPrice - paidAmount) * 100) / 100,
      paymentDate: paidAmount > 0 ? new Date(order.orderDate.getTime() + Math.random() * 7 * 24 * 60 * 60 * 1000) : undefined,
      status: order.paymentStatus === 'full-payment' ? 'paid-full' :
              order.paymentStatus === 'advance' ? 'paid-advance' :
              order.paymentStatus === 'bid' ? 'unpaid-bid' : 'unpaid-credit',
      paymentMethod: paidAmount > 0 ? ['Credit Card', 'Bank Transfer', 'Check', 'Cash'][Math.floor(Math.random() * 4)] : undefined
    };
  });
};

export const generateMockData = () => {
  const customers = generateCustomers();
  const employees = generateEmployees();
  const orders = generateOrders(customers, employees);
  const payments = generatePayments(orders);

  return { customers, employees, orders, payments };
};

export { generateId };