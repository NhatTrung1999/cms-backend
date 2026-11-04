import dayjs from 'dayjs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { buildQueryAutoSentCMS } from './cat9andcat12.helper';
import { addPortsToRows } from './shipment.helper';
import {
  ACTIVITY_TYPES as ACTIVITY_TYPES_CAT9,
  ActivityType as ActivityTypeCat9,
} from 'src/types/cat9andcat12';

/** A Cat9&12 row of the CMS query, after addPortsToRows (only the columns the mapper reads). */
export interface Cat9CmsRow {
  Date?: string | Date;
  Factory?: string;
  Factory_address?: string;
  Shipment_Date?: string | Date;
  Booking_No?: string;
  Customer_ID?: string;
  Invoice_Number?: string;
  Transport_Method: string;
  Port_Of_Arrival?: string;
  Port_Of_Departure?: string;
  Quantity?: number;
  Gross_Weight?: number | string;
}

/**
 * Turns one Cat9&12 row (already given its ports) into the CMS payload row(s).
 * It returns [] when `dockeyCMS` is not one of the activity types of the category.
 */
export const mapToCMSFormatCat9 = (
  item: Cat9CmsRow,
  dateFrom: string,
  dateTo: string,
  dockeyCMS: string,
) => {
  const date = item.Date;
  const factory = item.Factory;
  const factoryAddress = item.Factory_address;
  const shipmentDate = item.Shipment_Date;
  const bookingNo = item.Booking_No;
  const customerID = item.Customer_ID;
  const invoiceNumber = item.Invoice_Number;
  const transportMethod = item.Transport_Method; // AIR OR SEA
  const portOfArrival = item.Port_Of_Arrival; // DEPARTURE AND END PORT
  const portOfDeparture = item.Port_Of_Departure; // ST PORT
  const quantity = item.Quantity;
  const grossWeight = item.Gross_Weight; // ACTIVITY DATA
  const createdDate = dayjs().format('YYYY-MM-DD HH:mm:ss');

  let dockey = '';
  let transType = '';
  let portType = '';

  switch (transportMethod.trim().toLowerCase()) {
    case 'SEA'.trim().toLowerCase():
      dockey = '3.2.01';
      transType = '海運';
      portType = '海港';
      break;
    case 'AIR'.trim().toLowerCase():
      dockey = '3.2.02';
      transType = '空運';
      portType = '空港';
      break;
    case 'LAND'.trim().toLowerCase():
      dockey = '3.2.03';
      transType = '陸運';
      portType = '';
      break;
    default:
      dockey = '';
      break;
  }

  return ACTIVITY_TYPES_CAT9.filter((item) => item === dockeyCMS).map(
    (activityType: ActivityTypeCat9) => ({
      System: 'ERP',
      Corporation: 'LAI YIH',
      Factory: factory,
      Department: 'Shipping',
      DocKey: activityType.trim() === '5.3'.trim() ? '5.3' : dockey,
      SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
      EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
      ActivityType: activityType,
      DataType: activityType.trim() === '3.2' ? '1' : '999',
      DocType: 'Invoice',
      UndDoc: '銷貨單',
      DocFlow: '銷貨相關流程',
      DocDate: dayjs(date).format('YYYY/MM/DD'),
      DocDate2: dayjs(shipmentDate).format('YYYY/MM/DD'),
      DocNo: bookingNo,
      UndDocNo: '',
      CustVenName: customerID,
      InvoiceNo: invoiceNumber,
      TransType: transType,
      Departure: factoryAddress,
      Destination: portOfArrival,
      PortType: portType,
      StPort: portOfDeparture,
      ThPort: '',
      EndPort: portOfArrival,
      Product: '',
      Quity: quantity,
      Amount: '',
      ActivityData: +Number(grossWeight).toFixed(2) || 0,
      ActivityUnit: 'Kg',
      Unit: '',
      UnitWeight: '',
      Memo: '',
      CreateDateTime: createdDate,
      Creator: '',
      ActivitySource: 'Incineration (not energy recovery)',
    }),
  );
};

/**
 * CMS payload rows of one factory: reads its rows of the date range (ERP `db`), adds the ports (port of
 * arrival comes from `dbEIP`) and maps them with mapToCMSFormatCat9.
 */
export const getCat9AndCat12CMSRows = async (
  db: Sequelize,
  dbEIP: Sequelize,
  factory: string,
  dateFrom: string,
  dateTo: string,
  ry: string,
  dockeyCMS: string,
) => {
  const query = await buildQueryAutoSentCMS(
    dateFrom,
    dateTo,
    factory,
    ry,
    dbEIP,
  );

  const replacements =
    dateFrom && dateTo ? [dateFrom, dateTo, dateFrom, dateTo] : [];

  let data = await db.query<any>(query, {
    type: QueryTypes.SELECT,
    replacements,
  });

  data = await addPortsToRows(data, dbEIP);

  return data.flatMap((item) =>
    mapToCMSFormatCat9(item, dateFrom, dateTo, dockeyCMS),
  );
};
