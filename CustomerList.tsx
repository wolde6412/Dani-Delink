import React, { useState } from 'react';
import { Plus, Edit, Trash2, Mail, Phone } from 'lucide-react';
import { Customer } from '../../types';
import Table from '../Common/Table';
import Button from '../Common/Button';
import Modal from '../Common/Modal';
import CustomerForm from './CustomerForm';
import { format } from 'date-fns';

interface CustomerListProps {
  customers: Customer[];
  onAddCustomer: (customer: Omit<Customer, 'id' | 'createdAt'>) => void;
  onUpdateCustomer: (id: string, customer: Partial<Customer>) => void;
  onDeleteCustomer: (id: string) => void;
}

const CustomerList: React.FC<CustomerListProps> = ({
  customers,
  onAddCustomer,
  onUpdateCustomer,
  onDeleteCustomer
}) => {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingCustomer, setEditingCustomer] = useState<Customer | null>(null);

  const columns = [
    {
      key: 'id',
      label: 'Customer ID',
      sortable: true,
      render: (value: string) => (
        <span className="font-mono text-xs bg-gray-100 px-2 py-1 rounded">{value}</span>
      )
    },
    {
      key: 'name',
      label: 'Name',
      sortable: true,
      render: (value: string) => (
        <span className="font-medium text-gray-900">{value}</span>
      )
    },
    {
      key: 'companyName',
      label: 'Company',
      sortable: true
    },
    {
      key: 'contactNumber',
      label: 'Contact',
      render: (value: string) => (
        <div className="flex items-center space-x-1">
          <Phone className="w-3 h-3 text-gray-400" />
          <span className="text-sm">{value}</span>
        </div>
      )
    },
    {
      key: 'email',
      label: 'Email',
      render: (value: string) => (
        <div className="flex items-center space-x-1">
          <Mail className="w-3 h-3 text-gray-400" />
          <span className="text-sm">{value}</span>
        </div>
      )
    },
    {
      key: 'accountNumber',
      label: 'Account #',
      sortable: true
    },
    {
      key: 'createdAt',
      label: 'Created',
      sortable: true,
      render: (value: Date) => format(value, 'MMM dd, yyyy')
    },
    {
      key: 'actions',
      label: 'Actions',
      render: (_: any, customer: Customer) => (
        <div className="flex items-center space-x-2">
          <button
            onClick={(e) => {
              e.stopPropagation();
              setEditingCustomer(customer);
              setIsModalOpen(true);
            }}
            className="p-1 hover:bg-gray-100 rounded transition-colors"
          >
            <Edit className="w-4 h-4 text-gray-600" />
          </button>
          <button
            onClick={(e) => {
              e.stopPropagation();
              if (confirm('Are you sure you want to delete this customer?')) {
                onDeleteCustomer(customer.id);
              }
            }}
            className="p-1 hover:bg-red-100 rounded transition-colors"
          >
            <Trash2 className="w-4 h-4 text-red-600" />
          </button>
        </div>
      )
    }
  ];

  const handleSubmit = (customerData: Omit<Customer, 'id' | 'createdAt'>) => {
    if (editingCustomer) {
      onUpdateCustomer(editingCustomer.id, customerData);
    } else {
      onAddCustomer(customerData);
    }
    setIsModalOpen(false);
    setEditingCustomer(null);
  };

  const handleCloseModal = () => {
    setIsModalOpen(false);
    setEditingCustomer(null);
  };

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center">
        <div>
          <h2 className="text-xl font-semibold text-gray-900">Customer Management</h2>
          <p className="text-gray-600">Manage your customer database and contact information</p>
        </div>
        <Button
          onClick={() => setIsModalOpen(true)}
          icon={Plus}
        >
          Add Customer
        </Button>
      </div>

      <Table
        data={customers}
        columns={columns}
        searchPlaceholder="Search customers..."
      />

      <Modal
        isOpen={isModalOpen}
        onClose={handleCloseModal}
        title={editingCustomer ? 'Edit Customer' : 'Add New Customer'}
        size="lg"
      >
        <CustomerForm
          customer={editingCustomer}
          onSubmit={handleSubmit}
          onCancel={handleCloseModal}
        />
      </Modal>
    </div>
  );
};

export default CustomerList;