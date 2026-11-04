import type { Column } from 'exceljs';

/** Excel columns of the files exported by FilemanagementService (Cat6, Cat9&12, Cat5, Cat1&4, Cat7). */
export const CAT6_COLUMNS: Partial<Column>[] = [
  { header: 'Application Day', key: 'Application_Day' },
  { header: 'Document Number', key: 'Document_Number' },
  { header: 'Staff ID', key: 'Staff_ID' },
  { header: 'Dept', key: 'Dept' },
  { header: 'Round trip / One way', key: 'Round_trip_One_way' },
  { header: 'Start Time', key: 'Start_Time' },
  { header: 'End Time', key: 'End_Time' },
  { header: 'Business Trip Type', key: 'Business_Trip_Type' },
  { header: 'Departure', key: 'Departure' },
  { header: 'Destination', key: 'Destination' },
  { header: 'Transport', key: 'Transport' },
  {
    header: 'Number of nights stayed (Hotel)',
    key: 'Number_of_nights_stayed',
  },
];

export const CAT6_ACCOMMODATION_COLUMNS: Partial<Column>[] = [
  { header: 'Application Day', key: 'Application_Day' },
  { header: 'Document Number', key: 'Document_Number' },
  { header: 'Staff ID', key: 'Staff_ID' },
  { header: 'Dept', key: 'Dept' },
  { header: 'Business Trip Type', key: 'Business_Trip_Type' },
  {
    header: 'Number of nights stayed (Hotel)',
    key: 'Number_of_nights_stayed',
  },
];

export const CAT1_AND_CAT4_COLUMNS: Partial<Column>[] = [
  { header: 'No.', key: 'No' },
  { header: 'Factory Code', key: 'FactoryCode' },
  { header: 'Pur Date', key: 'PurDate' },
  { header: 'RK Date', key: 'RKDate' },
  { header: 'Purchase Order', key: 'PurNo' },
  { header: 'Received No.', key: 'ReceivedNo' },
  { header: 'Material No.', key: 'MatID' },
  { header: 'Qty.(Usage)', key: 'QtyUsage' },
  { header: 'Qty.(receive)', key: 'QtyReceive' },
  { header: 'Unit weight', key: 'UnitWeight' },
  { header: 'Weight (Unit：KG)', key: 'WeightUnitkg' },
  { header: 'Supplier Code', key: 'SupplierCode' },
  { header: 'Style', key: 'Style' },
  { header: 'Transportation Method', key: 'TransportationMethod' },
  { header: 'Departure', key: 'Departure' },
  {
    header: 'Third-country Land Transport (A)',
    key: 'ThirdCountryLandTransport',
  },
  { header: 'Port of Departure', key: 'PortOfDeparture' },
  { header: 'Port of Arrival', key: 'PortOfArrival' },
  {
    header: 'Factory (Domestic Land Transport)(B)',
    key: 'FactoryDomesticLandTransport',
  },
  { header: 'Destination', key: 'Destination' },
  {
    header: 'Land Transport Distance (A+B)',
    key: 'LandTransportDistance',
  },
  { header: 'Sea Transport Distance', key: 'SeaTransportDistance' },
  { header: 'Air Transport Distance', key: 'AirTransportDistance' },
  {
    header: 'Land Transport Ton-Kilometer',
    key: 'LandTransportTonKilometers',
  },
  {
    header: 'Sea Transport Ton-Kilometers',
    key: 'SeaTransportTonKilometers',
  },
  {
    header: 'Air Transport Ton-Kilomete',
    key: 'AirTransportTonKilometers',
  },
];

/** Same columns as the single-factory Cat7 export. */
export { CAT7_EXPORT_COLUMNS as CAT7_COLUMNS } from 'src/helper/cat7.columns';
export { CAT5_COLUMNS } from 'src/helper/cat5.columns';
export { CAT9_AND_CAT12_COLUMNS } from 'src/helper/cat9andcat12.columns';
