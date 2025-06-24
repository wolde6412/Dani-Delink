import React, { useState } from 'react';
import { Plus, Edit, Eye, Clock, CheckCircle, Package, Truck } from 'lucide-react';
import { Order, Customer, Employee } from '../../types';
import Table from '../Common/Table';
import Button from '../Common/Button';
import Modal from '../Common/Modal';
import OrderForm from './OrderForm';
import { format } from 'date-fns';

interface OrderListProps {
  orders: Order[];
  customers: Customer[];
  employees: Employee[];
  onAddOrder: (order: Omit<Order, 'id' | 'orderDate'>) => void;
  onUpdateOrder: (id: string, order: Partial<Order>) => void;
}

const OrderList: React.FC<OrderListProps> = ({
  orders,
  customers,
  employees,
  onAddOrder,
  onUpdateOrder
}) => {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingOrder, setEditingOrder] = useState<Order | null>(null);
  const [viewingOrder, setViewingOrder] = useState<Order | null>(null);

  const getStatusIcon = (status: Order['status']) => {
    switch (status) {
      case 'ordered': return <Clock className="w-4 h-4 text-yellow-600" />;
      case 'in-progress': return <Package className="w-4 h-4 text-blue-600" />;
      case 'on-schedule': return <CheckCircle className="w-4 h-4 text-green-600" />;
      case 'completed': return <CheckCircle className="w-4 h-4 text-green-700" />;
      case 'shipped': return <Truck className="w-4 h-4 text-purple-600" />;
      default: return null;
    }
  };

  const getStatusColor = (status: Order['status']) => {
    switch (status) {
      case 'ordered': return 'bg-yellow-100 text-yellow-800';
      case 'in-progress': return 'bg-blue-100 text-blue-800';
      case 'on-schedule': return 'bg-green-100 text-green-800';
      case 'completed': return 'bg-green-200 text-green-900';
      case 'shipped': return 'bg-purple-100 text-purple-800';
      default: return 'bg-gray-100 text-gray-800';
    }
  };

  const getPaymentStatusColor = (status: Order['paymentStatus']) => {
    switch (status) {
      case 'full-payment': return 'bg-green-100 text-green-800';
      case 'advance': return 'bg-blue-100 text-blue-800';
      case 'bid': return 'bg-yellow-100 text-yellow-800';
      case 'credit': return 'bg-orange-100 text-orange-800';
      default: return 'bg-gray-100 text-gray-800';
    }
  };

  const columns = [
    {
      key: 'id',
      label: 'Order ID',
      sortable: true,
      render: (value: string) => (
        <span className="font-mono text-xs bg-gray-100 px-2 py-1 rounded">{value}</span>
      )
    },
    {
      key: 'customerName',
      label: 'Customer',
      sortable: true,
      render: (value: string) => (
        <span className="font-medium text-gray-900">{value}</span>
      )
    },
    {
      key: 'orderType',
      label: 'Type',
      sortable: true,
      render: (value: string) => (
        <span className="text-sm font-medium text-gray-700">{value}</span>
      )
    },
    {
      key: 'quantity',
      label: 'Qty',
      sortable: true,
      render: (value: number) => (
        <span className="text-sm font-medium">{value.toLocaleString()}</span>
      )
    },
    {
      key: 'totalPrice',
      label: 'Total',
      sortable: true,
      render: (value: number) => (
        <span className="font-medium text-gray-900">${value.toFixed(2)}</span>
      )
    },
    {
      key: 'status',
      label: 'Status',
      sortable: true,
      render: (value: Order['status']) => (
        <div className="flex items-center space-x-2">
          {getStatusIcon(value)}
          <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${getStatusColor(value)}`}>
            {value.replace('-', ' ')}
          </span>
        </div>
      )
    },
    {
      key: 'paymentStatus',
      label: 'Payment',
      sortable: true,
      render: (value: Order['paymentStatus']) => (
        <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${getPaymentStatusColor(value)}`}>
          {value.replace('-', ' ')}
        </span>
      )
    },
    {
      key: 'assignedEmployeeName',
      label: 'Assigned To',
      sortable: true,
      render: (value: string) => (
        <span className="text-sm text-gray-700">{value}</span>
      )
    },
    {
      key: 'orderDate',
      label: 'Order Date',
      sortable: true,
      render: (value: Date) => format(value, 'MMM dd, yyyy')
    },
    {
      key: 'actions',
      label: 'Actions',
      render: (_: any, order: Order) => (
        <div className="flex items-center space-x-2">
          <button
            onClick={(e) => {
              e.stopPropagation();
              setViewingOrder(order);
            }}
            className="p-1 hover:bg-gray-100 rounded transition-colors"
            title="View Details"
          >
            <Eye className="w-4 h-4 text-gray-600" />
          </button>
          <button
            onClick={(e) => {
              e.stopPropagation();
              setEditingOrder(order);
              setIsModalOpen(true);
            }}
            className="p-1 hover:bg-gray-100 rounded transition-colors"
            title="Edit Order"
          >
            <Edit className="w-4 h-4 text-gray-600" />
          </button>
        </div>
      )
    }
  ];

  const handleSubmit = (orderData: Omit<Order, 'id' | 'orderDate'>) => {
    if (editingOrder) {
      onUpdateOrder(editingOrder.id, orderData);
    } else {
      onAddOrder(orderData);
    }
    setIsModalOpen(false);
    setEditingOrder(null);
  };

  const handleCloseModal = () => {
    setIsModalOpen(false);
    setEditingOrder(null);
  };

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center">
        <div>
          <h2 className="text-xl font-semibold text-gray-900">Order Management</h2>
          <p className="text-gray-600">Track and manage all customer orders and their status</p>
        </div>
        <Button
          onClick={() => setIsModalOpen(true)}
          icon={Plus}
        >
          New Order
        </Button>
      </div>

      <Table
        data={orders}
        columns={columns}
        searchPlaceholder="Search orders..."
      />

      <Modal
        isOpen={isModalOpen}
        onClose={handleCloseModal}
        title={editingOrder ? 'Edit Order' : 'Create New Order'}
        size="xl"
      >
        <OrderForm
          order={editingOrder}
          customers={customers}
          employees={employees}
          onSubmit={handleSubmit}
          onCancel={handleCloseModal}
        />
      </Modal>

      <Modal
        isOpen={!!viewingOrder}
        onClose={() => setViewingOrder(null)}
        title="Order Details"
        size="lg"
      >
        {viewingOrder && (
          <div className="space-y-6">
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-500">Order ID</label>
                <p className="mt-1 font-mono text-sm">{viewingOrder.id}</p>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-500">Order Date</label>
                <p className="mt-1">{format(viewingOrder.orderDate, 'PPP')}</p>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-500">Customer</label>
                <p className="mt-1 font-medium">{viewingOrder.customerName}</p>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-500">Assigned Employee</label>
                <p className="mt-1">{viewingOrder.assignedEmployeeName}</p>
              </div>
            </div>
            
            <div>
              <label className="block text-sm font-medium text-gray-500">Order Type</label>
              <p className="mt-1 font-medium">{viewingOrder.orderType}</p>
            </div>
            
            <div>
              <label className="block text-sm font-medium text-gray-500">Description</label>
              <p className="mt-1">{viewingOrder.description}</p>
            </div>
            
            <div className="grid grid-cols-3 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-500">Quantity</label>
                <p className="mt-1 font-medium">{viewingOrder.quantity.toLocaleString()}</p>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-500">Unit Price</label>
                <p className="mt-1 font-medium">${viewingOrder.unitPrice.toFixed(2)}</p>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-500">Total Price</label>
                <p className="mt-1 font-bold text-lg">${viewingOrder.totalPrice.toFixed(2)}</p>
              </div>
            </div>
            
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-500">Order Status</label>
                <div className="mt-1 flex items-center space-x-2">
                  {getStatusIcon(viewingOrder.status)}
                  <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${getStatusColor(viewingOrder.status)}`}>
                    {viewingOrder.status.replace('-', ' ')}
                  </span>
                </div>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-500">Payment Status</label>
                <span className={`mt-1 inline-flex px-2 py-1 text-xs font-semibold rounded-full ${getPaymentStatusColor(viewingOrder.paymentStatus)}`}>
                  {viewingOrder.paymentStatus.replace('-', ' ')}
                </span>
              </div>
            </div>
            
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-500">Estimated Delivery</label>
                <p className="mt-1">{format(viewingOrder.estimatedDelivery, 'PPP')}</p>
              </div>
              {viewingOrder.actualDelivery && (
                <div>
                  <label className="block text-sm font-medium text-gray-500">Actual Delivery</label>
                  <p className="mt-1">{format(viewingOrder.actualDelivery, 'PPP')}</p>
                </div>
              )}
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
};

export default OrderList;