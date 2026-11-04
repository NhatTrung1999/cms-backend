import type { Column } from 'exceljs';

/** Excel columns of the Cat5 (waste) export. */
export const CAT5_COLUMNS: Partial<Column>[] = [
  {
    header: 'Waste disposal date',
    key: 'Waste_disposal_date',
  },
  {
    header: 'Consolidated Waste',
    key: 'Consolidated_Waste',
  },
  {
    header: 'Waste Code',
    key: 'Waste_Code',
  },
  {
    header: 'Vendor Name',
    key: 'Vendor_Name',
  },
  {
    header: 'Vendor ID',
    key: 'Vendor_ID',
  },
  {
    header: 'Waste collection address',
    key: 'Waste_collection_address',
  },
  {
    header: 'Location Code',
    key: 'Location_Code',
  },
  {
    header: 'Transportation Distance (km)',
    key: 'Transportation_Distance_km',
  },
  {
    header: '*The type of waste',
    key: 'The_type_of_waste',
  },
  {
    header: '*Waste type',
    key: 'Waste_type',
  },
  {
    header: '*Waste Treatment method',
    key: 'Waste_Treatment_method',
  },
  {
    header: 'Treatment Method ID',
    key: 'Treatment_Method_ID',
  },
  {
    header: '*Weight of waste treated (Unit：kg)',
    key: 'Weight_of_waste_treated_Unit_kg',
  },
  {
    header: 'TKT (Ton-km)',
    key: 'TKT_Ton_km',
  },
];
