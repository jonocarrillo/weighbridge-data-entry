import { useState, useEffect } from 'react';
import WeighmasterTicket from '@/components/WeighmasterTicket';
import TicketList from '@/components/TicketList';
import ExportControls from '@/components/ExportControls';
import { WeighmasterEntry } from '@/types/weighmaster';
import { Scale, Plus, List } from 'lucide-react';

const STORAGE_KEY = 'weighmaster_entries';

const Index = () => {
  const [entries, setEntries] = useState<WeighmasterEntry[]>([]);
  const [editingEntry, setEditingEntry] = useState<WeighmasterEntry | null>(null);
  const [view, setView] = useState<'ticket' | 'list'>('ticket');

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
    setView('ticket');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handleDelete = (id: string) => {
    if (confirm('Delete this ticket?')) {
      setEntries(prev => prev.filter(e => e.id !== id));
    }
  };

  const handleNewTicket = () => {
    setEditingEntry(null);
    setView('ticket');
  };

  return (
    <div className="min-h-screen bg-background">
      {/* Header */}
      <header className="bg-card border-b-2 border-border">
        <div className="container mx-auto px-4 py-3">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-primary rounded flex items-center justify-center">
                <Scale className="w-6 h-6 text-primary-foreground" />
              </div>
              <div>
                <h1 className="text-lg font-bold">Weighmaster System</h1>
                <p className="text-xs text-muted-foreground">{entries.length} tickets saved</p>
              </div>
            </div>
            
            <div className="flex items-center gap-2">
              <button
                onClick={() => setView('ticket')}
                className={`px-3 py-1.5 rounded text-sm font-medium transition-colors ${
                  view === 'ticket' ? 'bg-primary text-primary-foreground' : 'bg-secondary'
                }`}
              >
                <Plus className="w-4 h-4 inline mr-1" />
                Ticket
              </button>
              <button
                onClick={() => setView('list')}
                className={`px-3 py-1.5 rounded text-sm font-medium transition-colors ${
                  view === 'list' ? 'bg-primary text-primary-foreground' : 'bg-secondary'
                }`}
              >
                <List className="w-4 h-4 inline mr-1" />
                List
              </button>
            </div>
          </div>
        </div>
      </header>

      <main className="container mx-auto px-4 py-6">
        {view === 'ticket' ? (
          <div className="flex flex-col items-center">
            <WeighmasterTicket
              onSubmit={handleSubmit}
              editEntry={editingEntry}
              onCancelEdit={() => setEditingEntry(null)}
            />
          </div>
        ) : (
          <div className="space-y-4">
            <div className="flex justify-between items-center">
              <h2 className="font-semibold">All Tickets</h2>
              <ExportControls entries={entries} />
            </div>
            <TicketList
              entries={entries}
              onEdit={handleEdit}
              onDelete={handleDelete}
              onNewTicket={handleNewTicket}
            />
          </div>
        )}
      </main>
    </div>
  );
};

export default Index;
