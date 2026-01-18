import { WeighmasterEntry } from '@/types/weighmaster';
import { Pencil, Trash2 } from 'lucide-react';

interface DataTableProps {
  entries: WeighmasterEntry[];
  onEdit: (entry: WeighmasterEntry) => void;
  onDelete: (id: string) => void;
}

const DataTable = ({ entries, onEdit, onDelete }: DataTableProps) => {
  if (entries.length === 0) {
    return (
      <div className="certificate-card rounded-lg p-8 text-center">
        <p className="text-muted-foreground">No entries yet. Add your first weighmaster certificate above.</p>
      </div>
    );
  }

  return (
    <div className="certificate-card rounded-lg overflow-hidden">
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="data-table-header">
              <th className="px-2 py-2 text-left">Invoice #</th>
              <th className="px-2 py-2 text-left">Date</th>
              <th className="px-2 py-2 text-left">Customer</th>
              <th className="px-2 py-2 text-left">Truck ID</th>
              <th className="px-2 py-2 text-left">Waste</th>
              <th className="px-2 py-2 text-left">Time In</th>
              <th className="px-2 py-2 text-left">Time Out</th>
              <th className="px-2 py-2 text-right">Weigh In</th>
              <th className="px-2 py-2 text-right">Weigh Out</th>
              <th className="px-2 py-2 text-right font-bold">Tonnage</th>
              <th className="px-2 py-2 text-right">Amount</th>
              <th className="px-2 py-2 text-left">Payment</th>
              <th className="px-2 py-2 text-center">Actions</th>
            </tr>
          </thead>
          <tbody className="font-mono text-xs">
            {entries.map((entry, index) => (
              <tr 
                key={entry.id} 
                className={`border-b border-border hover:bg-primary/10 transition-colors ${
                  index % 2 === 1 ? 'data-table-row-alt' : ''
                }`}
              >
                <td className="px-2 py-2 font-semibold">{entry.invoiceNumber}</td>
                <td className="px-2 py-2">{entry.date}</td>
                <td className="px-2 py-2 max-w-[120px] truncate" title={entry.customer}>
                  {entry.customer}
                </td>
                <td className="px-2 py-2">{entry.truckId}</td>
                <td className="px-2 py-2 max-w-[80px] truncate" title={entry.wasteData}>
                  {entry.wasteData}
                </td>
                <td className="px-2 py-2">{entry.timeIn}</td>
                <td className="px-2 py-2">{entry.timeOut}</td>
                <td className="px-2 py-2 text-right">{entry.weighIn.toLocaleString()}</td>
                <td className="px-2 py-2 text-right">{entry.weighOut.toLocaleString()}</td>
                <td className="px-2 py-2 text-right font-bold text-accent">{entry.tonnage.toFixed(2)}</td>
                <td className="px-2 py-2 text-right">${entry.amount.toFixed(2)}</td>
                <td className="px-2 py-2">
                  <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                    entry.paymentType === 'CASH' ? 'bg-green-100 text-green-800' :
                    entry.paymentType === 'CARD' ? 'bg-blue-100 text-blue-800' :
                    entry.paymentType === 'ACCOUNT' ? 'bg-amber-100 text-amber-800' :
                    'bg-gray-100 text-gray-600'
                  }`}>
                    {entry.paymentType || '-'}
                  </span>
                </td>
                <td className="px-2 py-2">
                  <div className="flex justify-center gap-1">
                    <button 
                      onClick={() => onEdit(entry)}
                      className="p-1 hover:bg-primary/20 rounded transition-colors"
                      title="Edit"
                    >
                      <Pencil className="w-3.5 h-3.5" />
                    </button>
                    <button 
                      onClick={() => onDelete(entry.id)}
                      className="p-1 hover:bg-destructive/20 rounded transition-colors text-destructive"
                      title="Delete"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default DataTable;
