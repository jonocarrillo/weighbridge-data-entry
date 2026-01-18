import { WeighmasterEntry } from '@/types/weighmaster';
import { Download, Copy, FileSpreadsheet } from 'lucide-react';
import { toast } from 'sonner';

interface ExportControlsProps {
  entries: WeighmasterEntry[];
}

const ExportControls = ({ entries }: ExportControlsProps) => {
  const generateCSV = (): string => {
    const headers = [
      'Invoice Number',
      'WASTE DATA',
      'TIME IN',
      'TIME OUT',
      'Memo',
      'weighin',
      'weighout',
      'TRUCKID',
      'Address',
      'TONAGE',
      'DATE',
      'CUSTOMER',
      'Weighmaster#',
      'Product/Service',
      'CITY',
      'AMOUNT',
      'PAYMENTYPE'
    ];

    const rows = entries.map(entry => [
      entry.invoiceNumber,
      entry.wasteData,
      entry.timeIn,
      entry.timeOut,
      entry.memo,
      entry.weighIn,
      entry.weighOut,
      entry.truckId,
      entry.address,
      entry.tonnage,
      entry.date,
      entry.customer,
      entry.weighmasterNumber,
      entry.productService,
      entry.city,
      entry.amount,
      entry.paymentType
    ]);

    const csvContent = [
      headers.join(','),
      ...rows.map(row => row.map(cell => `"${cell}"`).join(','))
    ].join('\n');

    return csvContent;
  };

  const handleDownloadCSV = () => {
    if (entries.length === 0) {
      toast.error('No entries to export');
      return;
    }
    
    const csv = generateCSV();
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `weighmaster_data_${new Date().toISOString().split('T')[0]}.csv`;
    link.click();
    toast.success('CSV downloaded successfully');
  };

  const handleCopyCSV = async () => {
    if (entries.length === 0) {
      toast.error('No entries to copy');
      return;
    }
    
    const csv = generateCSV();
    await navigator.clipboard.writeText(csv);
    toast.success('CSV copied to clipboard - paste into Excel for VBA');
  };

  const generateVBACode = (): string => {
    return `' VBA Import Code for Weighmaster Data
' Paste this into your Excel VBA module

Sub ImportWeighmasterCSV()
    Dim filePath As String
    Dim ws As Worksheet
    
    ' Open file dialog
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv")
    
    If filePath = "False" Then Exit Sub
    
    ' Create new worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Weighmaster_" & Format(Date, "YYYYMMDD")
    
    ' Import CSV
    With ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .Refresh BackgroundQuery:=False
    End With
    
    ' Format header row
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
    
    MsgBox "Import Complete!", vbInformation
End Sub`;
  };

  const handleCopyVBA = async () => {
    const vba = generateVBACode();
    await navigator.clipboard.writeText(vba);
    toast.success('VBA code copied to clipboard');
  };

  return (
    <div className="flex flex-wrap gap-2">
      <button onClick={handleDownloadCSV} className="btn-success flex items-center gap-2">
        <Download className="w-4 h-4" />
        Download CSV
      </button>
      <button onClick={handleCopyCSV} className="btn-secondary flex items-center gap-2">
        <Copy className="w-4 h-4" />
        Copy to Clipboard
      </button>
      <button onClick={handleCopyVBA} className="btn-secondary flex items-center gap-2">
        <FileSpreadsheet className="w-4 h-4" />
        Copy VBA Import Code
      </button>
    </div>
  );
};

export default ExportControls;
