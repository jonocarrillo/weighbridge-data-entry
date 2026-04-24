import { useState } from 'react';
import { WeighmasterEntry, createEmptyEntry, calculateTonnage } from '@/types/weighmaster';
import { Scale, Truck, FileText, DollarSign, Clock, MapPin } from 'lucide-react';

interface WeighmasterFormProps {
  onSubmit: (entry: WeighmasterEntry) => void;
  editEntry?: WeighmasterEntry | null;
  onCancelEdit?: () => void;
}

const WeighmasterForm = ({ onSubmit, editEntry, onCancelEdit }: WeighmasterFormProps) => {
  const [formData, setFormData] = useState(editEntry || createEmptyEntry());

  const handleChange = (field: string, value: string | number) => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const weighIn = formData.weighIn === '' || formData.weighIn === null || formData.weighIn === undefined ? 0 : Number(formData.weighIn);
    const weighOut = formData.weighOut === '' || formData.weighOut === null || formData.weighOut === undefined ? 0 : Number(formData.weighOut);
    const amount = formData.amount === '' || formData.amount === null || formData.amount === undefined ? 0 : Number(formData.amount);
    const tonnage = weighIn > 0 && weighOut > 0 ? calculateTonnage(weighIn, weighOut) : 0;
    const entry: WeighmasterEntry = {
      ...formData,
      id: editEntry?.id || crypto.randomUUID(),
      weighIn,
      weighOut,
      amount,
      tonnage,
    };
    onSubmit(entry);
    if (!editEntry) {
      setFormData(createEmptyEntry());
    }
  };

  const hasBothWeights = formData.weighIn !== '' && formData.weighIn !== 0 && formData.weighOut !== '' && formData.weighOut !== 0;
  const netWeight = hasBothWeights ? Math.abs(Number(formData.weighIn) - Number(formData.weighOut)) : 0;
  const calculatedTonnage = hasBothWeights ? calculateTonnage(Number(formData.weighIn), Number(formData.weighOut)) : 0;

  return (
    <form onSubmit={handleSubmit} className="certificate-card rounded-lg overflow-hidden">
      {/* Header */}
      <div className="certificate-header flex items-center gap-3">
        <Scale className="w-6 h-6" />
        <div>
          <h2 className="text-lg font-bold tracking-tight">WEIGHMASTER CERTIFICATE</h2>
          <p className="text-xs opacity-80">Data Entry Form</p>
        </div>
      </div>

      <div className="p-4 space-y-4">
        {/* Row 1: Invoice & Date */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          <div>
            <label className="form-label">Invoice Number</label>
            <input
              type="text"
              value={formData.invoiceNumber}
              onChange={(e) => handleChange('invoiceNumber', e.target.value)}
              className="input-field w-full"
              placeholder="103091"
            />
          </div>
          <div>
            <label className="form-label">Date</label>
            <input
              type="date"
              value={formData.date}
              onChange={(e) => handleChange('date', e.target.value)}
              className="input-field w-full"
            />
          </div>
          <div>
            <label className="form-label">Weighmaster #</label>
            <input
              type="text"
              value={formData.weighmasterNumber}
              onChange={(e) => handleChange('weighmasterNumber', e.target.value)}
              className="input-field w-full"
              placeholder="WM-001"
            />
          </div>
          <div>
            <label className="form-label">Truck ID</label>
            <div className="relative">
              <Truck className="absolute left-2 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
              <input
                type="text"
                value={formData.truckId}
                onChange={(e) => handleChange('truckId', e.target.value)}
                className="input-field w-full pl-8"
                placeholder="103091"
              />
            </div>
          </div>
        </div>

        {/* Row 2: Customer Info */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
          <div>
            <label className="form-label">Customer</label>
            <input
              type="text"
              value={formData.customer}
              onChange={(e) => handleChange('customer', e.target.value)}
              className="input-field w-full"
              placeholder="NATIONAL SCAN"
            />
          </div>
          <div>
            <label className="form-label">Address</label>
            <div className="relative">
              <MapPin className="absolute left-2 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
              <input
                type="text"
                value={formData.address}
                onChange={(e) => handleChange('address', e.target.value)}
                className="input-field w-full pl-8"
                placeholder="5626 Cherry Ave"
              />
            </div>
          </div>
          <div>
            <label className="form-label">City</label>
            <input
              type="text"
              value={formData.city}
              onChange={(e) => handleChange('city', e.target.value)}
              className="input-field w-full"
              placeholder="Long Beach, CA"
            />
          </div>
        </div>

        {/* Row 3: Time & Waste */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          <div>
            <label className="form-label flex items-center gap-1">
              <Clock className="w-3 h-3" /> Time In
            </label>
            <input
              type="time"
              value={formData.timeIn}
              onChange={(e) => handleChange('timeIn', e.target.value)}
              className="input-field w-full"
            />
          </div>
          <div>
            <label className="form-label flex items-center gap-1">
              <Clock className="w-3 h-3" /> Time Out
            </label>
            <input
              type="time"
              value={formData.timeOut}
              onChange={(e) => handleChange('timeOut', e.target.value)}
              className="input-field w-full"
            />
          </div>
          <div>
            <label className="form-label">Waste Data</label>
            <input
              type="text"
              value={formData.wasteData}
              onChange={(e) => handleChange('wasteData', e.target.value)}
              className="input-field w-full"
              placeholder="Type of waste"
            />
          </div>
          <div>
            <label className="form-label">Product/Service</label>
            <input
              type="text"
              value={formData.productService}
              onChange={(e) => handleChange('productService', e.target.value)}
              className="input-field w-full"
              placeholder="WEST"
            />
          </div>
        </div>

        {/* Row 4: Weights */}
        <div className="bg-secondary/50 rounded-lg p-3 border border-border">
          <div className="flex items-center gap-2 mb-3">
            <Scale className="w-4 h-4 text-accent" />
            <span className="font-semibold text-sm">Weight Data (lbs)</span>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            <div>
              <label className="form-label">Weigh In (Inbound)</label>
              <input
                type="number"
                value={formData.weighIn || ''}
                onChange={(e) => handleChange('weighIn', e.target.value)}
                className="input-field w-full text-right"
                placeholder="47820"
              />
            </div>
            <div>
              <label className="form-label">Weigh Out (Outbound)</label>
              <input
                type="number"
                value={formData.weighOut || ''}
                onChange={(e) => handleChange('weighOut', e.target.value)}
                className="input-field w-full text-right"
                placeholder="27940"
              />
            </div>
            <div>
              <label className="form-label">Net Weight (lbs)</label>
              <div className="input-field w-full text-right bg-muted font-bold">
                {netWeight.toLocaleString()}
              </div>
            </div>
            <div>
              <label className="form-label">Tonnage</label>
              <div className="input-field w-full text-right bg-primary/20 font-bold text-accent">
                {calculatedTonnage.toFixed(2)}
              </div>
            </div>
          </div>
        </div>

        {/* Row 5: Payment */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
          <div>
            <label className="form-label flex items-center gap-1">
              <DollarSign className="w-3 h-3" /> Amount
            </label>
            <input
              type="number"
              step="0.01"
              value={formData.amount || ''}
              onChange={(e) => handleChange('amount', e.target.value)}
              className="input-field w-full text-right"
              placeholder="1037.12"
            />
          </div>
          <div>
            <label className="form-label">Payment Type</label>
            <select
              value={formData.paymentType}
              onChange={(e) => handleChange('paymentType', e.target.value)}
              className="input-field w-full"
            >
              <option value="">Select...</option>
              <option value="CASH">CASH</option>
              <option value="CARD">CARD</option>
              <option value="ACCOUNT">ACCOUNT (AR)</option>
            </select>
          </div>
          <div>
            <label className="form-label flex items-center gap-1">
              <FileText className="w-3 h-3" /> Memo
            </label>
            <input
              type="text"
              value={formData.memo}
              onChange={(e) => handleChange('memo', e.target.value)}
              className="input-field w-full"
              placeholder="Special notes..."
            />
          </div>
        </div>

        {/* Actions */}
        <div className="flex gap-3 pt-2">
          <button type="submit" className="btn-primary flex-1">
            {editEntry ? 'Update Entry' : 'Add Entry'}
          </button>
          {editEntry && onCancelEdit && (
            <button type="button" onClick={onCancelEdit} className="btn-secondary">
              Cancel
            </button>
          )}
        </div>
      </div>
    </form>
  );
};

export default WeighmasterForm;
