import React, { useState, useEffect } from 'react';
import { Order, Customer, Employee } from '../../types';
import Button from '../Common/Button';
import { generateId } from '../../utils/dataGenerator';

interface OrderFormProps {
  order?: Order | null;
  customers: Customer[];
  employees: Employee[];
  onSubmit: (order: Omit<Order, 'id' | 'orderDate'>) => void;
  onCancel: () => void;
}

const OrderForm: React.FC<OrderFormProps> = ({
  order,
  customers,
  employees,
  onSubmit,
  onCancel
}) => {
  const [formData, setFormData] = useState({
    customerId: '',
    customerName: '',
    orderType: '',
    description: '',
    quantity: 1,
    unitPrice: 0,
    totalPrice: 0,
    status: 'ordered' as Order['status'],
    paymentStatus: 'bid' as Order['paymentStatus'],
    assignedEmployeeId: '',
    assignedEmployeeName: '',
    estimatedDelivery: '',
    actualDelivery: ''
  });

  useEffect(() => {
    if (order) {
      setFormData({
        customerId: order.customerId,
        customerName: order.customerName,
        orderType: order.orderType,
        description: order.description,
        quantity: order.quantity,
        unitPrice: order.unitPrice,
        totalPrice: order.totalPrice,
        status: order.status,
        paymentStatus: order.paymentStatus,
        assignedEmployeeId: order.assignedEmployeeId,
        assignedEmployeeName: order.assignedEmployeeName,
        estimatedDelivery: order.estimatedDelivery.toISOString().split('T')[0],
        actualDelivery: order.actualDelivery ? order.actualDelivery.toISOString().split('T')[0] : ''
      });
    }
  }, [order]);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const orderData = {
      ...formData,
      totalPrice: formData.quantity * formData.unitPrice,
      estimatedDelivery: new Date(formData.estimatedDelivery),
      actualDelivery: formData.actualDelivery ? new Date(formData.actualDelivery) : undefined
    };
    onSubmit(orderData);
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    
    if (name === 'customerId') {
      const selectedCustomer = customers.find(c => c.id === value);
      setFormData({
        ...formData,
        customerId: value,
        customerName: selectedCustomer?.name || ''
      });
    } else if (name === 'assignedEmployeeId') {
      const selectedEmployee = employees.find(e => e.id === value);
      setFormData({
        ...formData,
        assignedEmployeeId: value,
        assignedEmployeeName: selectedEmployee?.name || ''
      });
    } else if (name === 'quantity' || name === 'unitPrice') {
      const numValue = parseFloat(value) || 0;
      const updatedFormData = { ...formData, [name]: numValue };
      updatedFormData.totalPrice = updatedFormData.quantity * updatedFormData.unitPrice;
      setFormData(updatedFormData);
    } else {
      setFormData({
        ...formData,
        [name]: value
      });
    }
  };

  return (
    <form onSubmit={handleSubmit} className="space-y-6">
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Customer *
          </label>
          <select
            name="customerId"
            value={formData.customerId}
            onChange={handleChange}
            required
            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          >
            <option value="">Select a customer</option>
            {customers.map(customer => (
              <option key={customer.id} value={customer.id}>
                {customer.name} - {customer.companyName}
              </option>
            ))}
          </select>
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Assigned Employee *
          </label>
          <select
            name="assignedEmployeeId"
            value={formData.assignedEmployeeId}
            onChange={handleChange}
            required
            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          >
            <option value="">Select an employee</option>
            {employees.map(employee => (
              <option key={employee.id} value={employee.id}>
                {employee.name} - {employee.role}
              </option>
            ))}
          </select>
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Order Type *
          </label>
          <select
            name="orderType"
            value={formData.orderType}
            onChange={handleChange}
            required
            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          >
            <option value="">Select order type</option>
            <option value="Book Publishing">Book Publishing</option>
            <option value="Magazine Layout">Magazine Layout</option>
            <option value="Brochure Design">Brochure Design</option>
            <option value="Annual Report">Annual Report</option>
            <option value="Marketing Materials">Marketing Materials</option>
            <option value="Corporate Newsletter">Corporate Newsletter</option>
            <option value="Product Catalog">Product Catalog</option>
            <option value="Training Manual">Training Manual</option>
          </select>
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Status
          </label>
          <select
            name="status"
            value={formData.status}
            onChange={handleChange}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          >
            <option value="ordered">Ordered</option>
            <option value="in-progress">In Progress</option>
            <option value="on-schedule">On Schedule</option>
            <option value="completed">Completed</option>
            <option value="shipped">Shipped</option>
          </select>
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Payment Status
          </label>
          <select
            name="paymentStatus"
            value={formData.paymentStatus}
            onChange={handleChange}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          >
            <option value="bid">Bid</option>
            <option value="advance">Advance</option>
            <option value="full-payment">Full Payment</option>
            <option value="credit">Credit</option>
          </select>
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Quantity *
          </label>
          <input
            type="number"
            name="quantity"
            value={formData.quantity}
            onChange={handleChange}
            min="1"
            required
            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          />
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Unit Price *
          </label>
          <input
            type="number"
            name="unitPrice"
            value={formData.unitPrice}
            onChange={handleChange}
            min="0"
            step="0.01"
            required
            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          />
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Total Price
          </label>
          <input
            type="text"
            value={`$${formData.totalPrice.toFixed(2)}`}
            readOnly
            className="w-full px-3 py-2 border border-gray-300 rounded-lg bg-gray-50 font-medium"
          />
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Estimated Delivery *
          </label>
          <input
            type="date"
            name="estimatedDelivery"
            value={formData.estimatedDelivery}
            onChange={handleChange}
            required
            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          />
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Actual Delivery
          </label>
          <input
            type="date"
            name="actualDelivery"
            value={formData.actualDelivery}
            onChange={handleChange}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          />
        </div>
      </div>

      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">
          Description *
        </label>
        <textarea
          name="description"
          value={formData.description}
          onChange={handleChange}
          required
          rows={3}
          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          placeholder="Describe the order requirements and specifications"
        />
      </div>

      <div className="flex justify-end space-x-3 pt-4 border-t border-gray-200">
        <Button
          type="button"
          variant="outline"
          onClick={onCancel}
        >
          Cancel
        </Button>
        <Button type="submit">
          {order ? 'Update Order' : 'Create Order'}
        </Button>
      </div>
    </form>
  );
};

export default OrderForm;