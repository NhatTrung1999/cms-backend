import type { Column } from 'exceljs';

/** Excel columns of the Cat9&12 export (used by the single-factory and the "ALL" export). */
export const CAT9_AND_CAT12_COLUMNS: Partial<Column>[] = [
  {
    header: 'No.',
    key: 'No',
  },
  {
    header: 'Date',
    key: 'Date',
  },
  {
    header: 'RY Product',
    key: 'RYProduct',
  },
  {
    header: 'RY Component',
    key: 'RYComponent',
  },
  {
    header: 'Component Name',
    key: 'ComponentName',
  },
  {
    header: 'Shipment Date',
    key: 'Shipment_Date',
  },
  {
    header: 'Invoice Number',
    key: 'Invoice_Number',
  },
  {
    header: 'Article Name',
    key: 'Article_Name',
  },
  {
    header: 'Article ID',
    key: 'Article_ID',
  },
  {
    header: 'Quantity',
    key: 'Quantity',
  },
  {
    header: 'Gross Weight',
    key: 'Gross_Weight',
  },
  {
    header: 'Customer ID',
    key: 'Customer_ID',
  },
  {
    header: 'Local land transportation',
    key: 'Local_Land_Transportation',
  },
  {
    header: 'Port of Departure',
    key: 'Port_Of_Departure',
  },
  {
    header: 'Port of Arrival',
    key: 'Port_Of_Arrival',
  },
  {
    header: 'Land Transport Distance',
    key: 'Land_Transport_Distance',
  },
  {
    header: 'Sea Transport Distance',
    key: 'Sea_Transport_Distance',
  },
  {
    header: 'Air Transport Distance',
    key: 'Air_Transport_Distance',
  },
  {
    header: 'Transport Method',
    key: 'Transport_Method',
  },
  {
    header: 'Land Transport Ton-Kilometers',
    key: 'Land_Transport_Ton_Kilometers',
  },
  {
    header: 'Sea Transport Ton-Kilometers',
    key: 'Sea_Transport_Ton_Kilometers',
  },
  {
    header: 'Air Transport Ton-Kilometers',
    key: 'Air_Transport_Ton_Kilometers',
  },
];
