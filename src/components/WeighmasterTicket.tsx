import { useState, useEffect } from 'react';
import { WeighmasterEntry, createEmptyEntry, calculateTonnage } from '@/types/weighmaster';

interface WeighmasterTicketProps {
  onSubmit: (entry: WeighmasterEntry) => void;
  editEntry?: WeighmasterEntry | null;
  onCancelEdit?: () => void;
}

const WeighmasterTicket = ({ onSubmit, editEntry, onCancelEdit }: WeighmasterTicketProps) => {
  const [formData, setFormData] = useState(editEntry || createEmptyEntry());

  useEffect(() => {
    if (editEntry) {
      setFormData(editEntry);
    } else {
      setFormData(createEmptyEntry());
    }
  }, [editEntry]);

  const handleChange = (field: string, value: string | number) => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const tonnage = calculateTonnage(Number(formData.weighIn), Number(formData.weighOut));
    const entry: WeighmasterEntry = {
      ...formData,
      id: editEntry?.id || crypto.randomUUID(),
      weighIn: Number(formData.weighIn),
      weighOut: Number(formData.weighOut),
      amount: Number(formData.amount),
      tonnage,
    };
    onSubmit(entry);
    if (!editEntry) {
      setFormData(createEmptyEntry());
    }
  };

  const netWeight = Math.abs(Number(formData.weighIn) - Number(formData.weighOut));
  const calculatedTonnage = calculateTonnage(Number(formData.weighIn), Number(formData.weighOut));

  return (
    <form onSubmit={handleSubmit} className="ticket-container">
      {/* Ticket Header */}
      <div className="ticket-header">
        <div className="ticket-header-text">WEIGHMASTER CERTIFICATE</div>
        <p className="ticket-subtext">
          THIS IS TO CERTIFY that the following described commodity was weighed, measured or counted
          by a weighmaster whose signature is on this certificate.
        </p>
      </div>

      {/* Weighed At & Ticket Number */}
      <div className="ticket-row">
        <div className="ticket-field flex-1">
          <span className="ticket-label">WEIGHED AT</span>
          <input
            type="text"
            value={formData.address}
            onChange={(e) => handleChange('address', e.target.value)}
            className="ticket-input"
            placeholder="Location address"
          />
        </div>
        <div className="ticket-field w-28">
          <span className="ticket-label-right">Ticket #</span>
          <input
            type="text"
            value={formData.invoiceNumber}
            onChange={(e) => handleChange('invoiceNumber', e.target.value)}
            className="ticket-input text-right font-bold text-lg"
            placeholder="103091"
          />
        </div>
      </div>

      {/* Customer Info */}
      <div className="ticket-section">
        <div className="ticket-field">
          <span className="ticket-label">Weighed For/Seller:</span>
          <input
            type="text"
            value={formData.customer}
            onChange={(e) => handleChange('customer', e.target.value)}
            className="ticket-input"
            placeholder="NATIONAL SCAN"
          />
        </div>

        <div className="ticket-field">
          <span className="ticket-label">Company Name/Hauler:</span>
          <input
            type="text"
            value={formData.wasteData}
            onChange={(e) => handleChange('wasteData', e.target.value)}
            className="ticket-input"
            placeholder="Company or hauler name"
          />
        </div>

        <div className="ticket-field">
          <span className="ticket-label">Waste Origin:</span>
          <input
            type="text"
            value={formData.city}
            onChange={(e) => handleChange('city', e.target.value)}
            className="ticket-input"
            placeholder="L.A."
          />
        </div>

        <div className="ticket-field">
          <span className="ticket-label">Product Code:</span>
          <input
            type="text"
            value={formData.productService}
            onChange={(e) => handleChange('productService', e.target.value)}
            className="ticket-input"
            placeholder="WEST"
          />
        </div>

        <div className="ticket-row">
          <div className="ticket-field flex-1">
            <span className="ticket-label">Truck #</span>
            <input
              type="text"
              value={formData.truckId}
              onChange={(e) => handleChange('truckId', e.target.value)}
              className="ticket-input"
              placeholder="103091"
            />
          </div>
          <div className="ticket-field flex-1">
            <span className="ticket-label">Weighmaster #</span>
            <div className="ticket-input bg-muted font-bold">{formData.invoiceNumber || '------'}</div>
          </div>
        </div>

        <div className="ticket-field">
          <span className="ticket-label">Special Notes:</span>
          <input
            type="text"
            value={formData.memo}
            onChange={(e) => handleChange('memo', e.target.value)}
            className="ticket-input"
            placeholder="700+40 thang 51 sun togi"
          />
        </div>
      </div>

      {/* Payment Section */}
      <div className="ticket-payment-section">
        <div className="ticket-row items-center">
          <span className="ticket-label mr-2">PAID:</span>
          <label className="ticket-checkbox-label">
            <input
              type="radio"
              name="paymentType"
              value="CARD"
              checked={formData.paymentType === 'CARD'}
              onChange={(e) => handleChange('paymentType', e.target.value)}
              className="ticket-radio"
            />
            CARD
          </label>
          <label className="ticket-checkbox-label">
            <input
              type="radio"
              name="paymentType"
              value="CASH"
              checked={formData.paymentType === 'CASH'}
              onChange={(e) => handleChange('paymentType', e.target.value)}
              className="ticket-radio"
            />
            CASH
          </label>
          <label className="ticket-checkbox-label">
            <input
              type="radio"
              name="paymentType"
              value="ACCOUNT"
              checked={formData.paymentType === 'ACCOUNT'}
              onChange={(e) => handleChange('paymentType', e.target.value)}
              className="ticket-radio"
            />
            ACCOUNT
          </label>
          <div className="flex-1" />
          <span className="ticket-label">$</span>
          <input
            type="number"
            step="0.01"
            value={formData.amount || ''}
            onChange={(e) => handleChange('amount', e.target.value)}
            className="ticket-input w-28 text-right font-bold"
            placeholder="1037.12"
          />
        </div>
      </div>

      {/* Time Stamps */}
      <div className="ticket-timestamp-section">
        <div className="ticket-timestamp-row">
          <input
            type="time"
            value={formData.timeIn}
            onChange={(e) => handleChange('timeIn', e.target.value)}
            className="ticket-input-small"
          />
          <input
            type="date"
            value={formData.date}
            onChange={(e) => handleChange('date', e.target.value)}
            className="ticket-input-small"
          />
          <span className="ticket-timestamp-label">INBOUND</span>
        </div>
        <div className="ticket-timestamp-detail">
          Truck ID: {formData.truckId || '------'}
          <br />
          Inbound Wt: <span className="font-bold">{Number(formData.weighIn).toLocaleString()} lb</span>
        </div>

        <div className="ticket-divider" />

        <div className="ticket-timestamp-row">
          <input
            type="time"
            value={formData.timeOut}
            onChange={(e) => handleChange('timeOut', e.target.value)}
            className="ticket-input-small"
          />
          <span className="ticket-timestamp-label">OUTBOUND</span>
        </div>
        <div className="ticket-timestamp-detail">
          Truck ID: {formData.truckId || '------'}
        </div>
      </div>

      {/* Weight Summary */}
      <div className="ticket-weight-section">
        <div className="ticket-weight-row">
          <span>Outbound Weight:</span>
          <div className="flex items-center gap-2">
            <input
              type="number"
              value={formData.weighOut || ''}
              onChange={(e) => handleChange('weighOut', e.target.value)}
              className="ticket-input-weight"
              placeholder="27940"
            />
            <span className="text-xs">lb</span>
          </div>
        </div>
        <div className="ticket-weight-row">
          <span>Inbound Weight:</span>
          <div className="flex items-center gap-2">
            <input
              type="number"
              value={formData.weighIn || ''}
              onChange={(e) => handleChange('weighIn', e.target.value)}
              className="ticket-input-weight"
              placeholder="47820"
            />
            <span className="text-xs">lb</span>
          </div>
        </div>
        <div className="ticket-weight-row text-muted-foreground">
          <span>ED</span>
          <span></span>
        </div>
        <div className="ticket-weight-row font-bold text-lg border-t border-border pt-2 mt-2">
          <span>Net Weight:</span>
          <span>{netWeight.toLocaleString()} lb</span>
        </div>
      </div>

      {/* Footer */}
      <div className="ticket-footer">
        <div className="ticket-row">
          <div className="ticket-field flex-1">
            <span className="ticket-label">Net Weight</span>
            <div className="ticket-result">{calculatedTonnage.toFixed(2)}</div>
          </div>
          <div className="ticket-field text-center">
            <span className="ticket-label">Tons</span>
          </div>
          <div className="ticket-field flex-1 text-right">
            <span className="ticket-label">-Weighmaster</span>
          </div>
        </div>
        <div className="ticket-signature">
          BY: ____________________________
        </div>
      </div>

      {/* Actions */}
      <div className="ticket-actions">
        <button type="submit" className="btn-primary w-full">
          {editEntry ? 'Update Ticket' : 'Save Ticket'}
        </button>
        {editEntry && onCancelEdit && (
          <button type="button" onClick={onCancelEdit} className="btn-secondary w-full mt-2">
            Cancel
          </button>
        )}
      </div>
    </form>
  );
};

export default WeighmasterTicket;
