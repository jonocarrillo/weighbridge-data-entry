import { WeighmasterEntry } from '@/types/weighmaster';
import { Pencil, Trash2, Plus } from 'lucide-react';

interface TicketListProps {
  entries: WeighmasterEntry[];
  onEdit: (entry: WeighmasterEntry) => void;
  onDelete: (id: string) => void;
  onNewTicket: () => void;
}

const TicketList = ({ entries, onEdit, onDelete, onNewTicket }: TicketListProps) => {
  if (entries.length === 0) {
    return (
      <div className="ticket-container text-center py-12">
        <p className="text-muted-foreground mb-4">No tickets yet</p>
        <button onClick={onNewTicket} className="btn-primary">
          <Plus className="w-4 h-4 inline mr-2" />
          Create First Ticket
        </button>
      </div>
    );
  }

  return (
    <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-3">
      {entries.map((entry) => (
        <div key={entry.id} className="ticket-mini">
          <div className="ticket-mini-header">
            <span className="font-bold">#{entry.invoiceNumber}</span>
            <span className="text-xs">{entry.date}</span>
          </div>
          
          <div className="ticket-mini-body">
            <div className="ticket-mini-row">
              <span className="text-muted-foreground">Customer:</span>
              <span className="font-medium truncate">{entry.customer || '-'}</span>
            </div>
            <div className="ticket-mini-row">
              <span className="text-muted-foreground">Truck:</span>
              <span className="font-mono">{entry.truckId || '-'}</span>
            </div>
            <div className="ticket-mini-row">
              <span className="text-muted-foreground">Net Weight:</span>
              <span className="font-mono">
                {Math.abs(entry.weighIn - entry.weighOut).toLocaleString()} lb
              </span>
            </div>
            <div className="ticket-mini-row">
              <span className="text-muted-foreground">Tonnage:</span>
              <span className="font-bold text-accent">{entry.tonnage.toFixed(2)} T</span>
            </div>
            <div className="ticket-mini-row">
              <span className="text-muted-foreground">Amount:</span>
              <span className="font-bold">${entry.amount.toFixed(2)}</span>
            </div>
            <div className="ticket-mini-row">
              <span className="text-muted-foreground">Payment:</span>
              <span className={`ticket-payment-badge ${
                entry.paymentType === 'CASH' ? 'badge-cash' :
                entry.paymentType === 'CARD' ? 'badge-card' :
                entry.paymentType === 'ACCOUNT' ? 'badge-account' : ''
              }`}>
                {entry.paymentType || '-'}
              </span>
            </div>
          </div>

          <div className="ticket-mini-footer">
            <button onClick={() => onEdit(entry)} className="ticket-mini-btn">
              <Pencil className="w-4 h-4" /> Edit
            </button>
            <button onClick={() => onDelete(entry.id)} className="ticket-mini-btn text-destructive">
              <Trash2 className="w-4 h-4" /> Delete
            </button>
          </div>
        </div>
      ))}
    </div>
  );
};

export default TicketList;
