export interface WeighmasterEntry {
  id: string;
  invoiceNumber: string;
  wasteData: string;
  timeIn: string;
  timeOut: string;
  memo: string;
  weighIn: number;
  weighOut: number;
  truckId: string;
  address: string;
  tonnage: number;
  date: string;
  customer: string;
  weighmasterNumber: string;
  productService: string;
  city: string;
  amount: number;
  paymentType: 'CASH' | 'CARD' | 'ACCOUNT' | '';
}

export const createEmptyEntry = (): Omit<WeighmasterEntry, 'id' | 'tonnage'> => ({
  invoiceNumber: '',
  wasteData: '',
  timeIn: '',
  timeOut: '',
  memo: '',
  weighIn: 0,
  weighOut: 0,
  truckId: '',
  address: '',
  date: new Date().toISOString().split('T')[0],
  customer: '',
  weighmasterNumber: '',
  productService: '',
  city: '',
  amount: 0,
  paymentType: '',
});

export const calculateTonnage = (weighIn: number, weighOut: number): number => {
  const netWeight = Math.abs(weighIn - weighOut);
  return Number((netWeight / 2000).toFixed(2));
};
