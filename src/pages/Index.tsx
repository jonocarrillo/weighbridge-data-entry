import { useState, useEffect } from 'react';
import WeighmasterForm from '@/components/WeighmasterForm';
import DataTable from '@/components/DataTable';
import ExportControls from '@/components/ExportControls';
import { WeighmasterEntry } from '@/types/weighmaster';
import { Scale, Database } from 'lucide-react';

const STORAGE_KEY = 'weighmaster_entries';

const Index = () => {
  const [entries, setEntries] = useState<WeighmasterEntry[]>([]);
  const [editingEntry, setEditingEntry] = useState<WeighmasterEntry | null>(null);

  // Load from localStorage on mount
  useEffect(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        setEntries(JSON.parse(saved));
      } catch (e) {
        console.error('Failed to parse saved entries:', e);
      }
    }
  }, []);

  // Save to localStorage on change
  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
  }, [entries]);

  const handleSubmit = (entry: WeighmasterEntry) => {
    if (editingEntry) {
      setEntries(prev => prev.map(e => e.id === entry.id ? entry : e));
      setEditingEntry(null);
    } else {
      setEntries(prev => [entry, ...prev]);
    }
  };

  const handleEdit = (entry: WeighmasterEntry) => {
    setEditingEntry(entry);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handleDelete = (id: string) => {
    if (confirm('Are you sure you want to delete this entry?')) {
      setEntries(prev => prev.filter(e => e.id !== id));
    }
  };

  const handleCancelEdit = () => {
    setEditingEntry(null);
  };

  const totalTonnage = entries.reduce((sum, e) => sum + e.tonnage, 0);
  const totalAmount = entries.reduce((sum, e) => sum + e.amount, 0);

  return (
    <div className="min-h-screen bg-background">
      {/* Header */}
      <header className="bg-card border-b border-border sticky top-0 z-10">
        <div className="container mx-auto px-4 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-primary rounded-lg flex items-center justify-center">
                <Scale className="w-6 h-6 text-primary-foreground" />
              </div>
              <div>
                <h1 className="text-xl font-bold text-foreground">Weighmaster Data Entry</h1>
                <p className="text-xs text-muted-foreground">VBA-Compatible Export System</p>
              </div>
            </div>
            <div className="flex items-center gap-4 text-sm">
              <div className="flex items-center gap-2 px-3 py-1.5 bg-secondary rounded-md">
                <Database className="w-4 h-4 text-muted-foreground" />
                <span className="font-mono font-bold">{entries.length}</span>
                <span className="text-muted-foreground">entries</span>
              </div>
            </div>
          </div>
        </div>
      </header>

      <main className="container mx-auto px-4 py-6 space-y-6">
        {/* Summary Stats */}
        {entries.length > 0 && (
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            <div className="certificate-card rounded-lg p-4">
              <p className="text-xs text-muted-foreground uppercase tracking-wide">Total Entries</p>
              <p className="text-2xl font-bold font-mono">{entries.length}</p>
            </div>
            <div className="certificate-card rounded-lg p-4">
              <p className="text-xs text-muted-foreground uppercase tracking-wide">Total Tonnage</p>
              <p className="text-2xl font-bold font-mono text-accent">{totalTonnage.toFixed(2)}</p>
            </div>
            <div className="certificate-card rounded-lg p-4">
              <p className="text-xs text-muted-foreground uppercase tracking-wide">Total Amount</p>
              <p className="text-2xl font-bold font-mono">${totalAmount.toFixed(2)}</p>
            </div>
            <div className="certificate-card rounded-lg p-4">
              <p className="text-xs text-muted-foreground uppercase tracking-wide">Avg Per Entry</p>
              <p className="text-2xl font-bold font-mono">
                ${entries.length > 0 ? (totalAmount / entries.length).toFixed(2) : '0.00'}
              </p>
            </div>
          </div>
        )}

        {/* Form */}
        <WeighmasterForm 
          onSubmit={handleSubmit} 
          editEntry={editingEntry}
          onCancelEdit={handleCancelEdit}
        />

        {/* Export Controls */}
        <div className="flex items-center justify-between flex-wrap gap-4">
          <h2 className="text-lg font-semibold flex items-center gap-2">
            <Database className="w-5 h-5" />
            Data Records
          </h2>
          <ExportControls entries={entries} />
        </div>

        {/* Data Table */}
        <DataTable 
          entries={entries} 
          onEdit={handleEdit} 
          onDelete={handleDelete} 
        />
      </main>

      {/* Footer */}
      <footer className="border-t border-border mt-12">
        <div className="container mx-auto px-4 py-4 text-center text-xs text-muted-foreground">
          Weighmaster Certificate Data Entry System • Export to CSV for VBA/Excel Import
        </div>
      </footer>
    </div>
  );
};

export default Index;
