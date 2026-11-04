import type { Column } from 'exceljs';

/** Excel columns of the Cat7 export (one factory or all factories). */
export const CAT7_EXPORT_COLUMNS: Partial<Column>[] = [
  { header: 'Staff ID', key: 'Staff_ID' },
  { header: 'Residential address', key: 'Residential_address' },
  { header: 'Main transportation type', key: 'Main_transportation_type' },
  { header: 'km', key: 'km' },
  { header: 'Number of working days', key: 'Number_of_working_days' },
  { header: 'PKT (p-km)', key: 'PKT_p_km' },
];
