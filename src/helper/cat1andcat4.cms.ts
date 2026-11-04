import dayjs from 'dayjs';
import { ACTIVITY_TYPES, ActivityType } from 'src/types/cat1andcat4';

/** A row of the Cat1&4 CMS query (only the columns the mapper reads). */
export interface Cat1And4CmsRow {
  FactoryCode?: string;
  MatID?: string;
  RKDate?: string | Date;
  SupplierCode?: string;
  TransportationMethod?: string;
  Departure?: string;
  PortOfDeparture?: string;
  PortOfArrival?: string;
  Destination?: string;
  WeightUnitkg?: number;
  QtyReceive?: number;
  ReceivedNo?: string;
}

/**
 * Turns one Cat1&4 row read from a factory database into the CMS payload row(s).
 * It returns [] when `dockeyCMS` is not one of the activity types of the category.
 */
export const mapToCMSFormatCat1And4 = (
  item: Cat1And4CmsRow,
  dateFrom: string,
  dateTo: string,
  dockeyCMS: string,
) => {
  const factory = item.FactoryCode ?? '';
  const matId = item.MatID ?? '';
  const docKey = matId.substring(0, 1);
  if (docKey.startsWith('W')) {
    return [];
  }
  const docDate = item.RKDate ? dayjs(item.RKDate).format('YYYY/MM/DD') : '';
  const custVenName = item.SupplierCode ?? '';
  const transportationMethod = item.TransportationMethod ?? '';
  const departure = item.Departure ?? '';
  const portOfDeparture = item.PortOfDeparture ?? '';
  const portOfArrival = item.PortOfArrival ?? '';
  const destination = item.Destination ?? '';
  const activityData = item.WeightUnitkg ?? 0;
  const qtyReceive = item.QtyReceive ?? 0;
  const receivedNo = item.ReceivedNo ?? '';

  let transType = '';
  let portType = '';

  switch (transportationMethod.trim().toLowerCase()) {
    case 'SEA'.trim().toLowerCase():
    case 'SEA+LAND'.trim().toLowerCase():
      transType = '海運';
      portType = '海港';
      break;
    case 'AIR'.trim().toLowerCase():
      transType = '空運';
      portType = '空港';
      break;
    case 'LAND'.trim().toLowerCase():
      transType = '陸運';
      portType = '';
      break;
    default:
      break;
  }

  return ACTIVITY_TYPES.filter((item) => item === dockeyCMS).map(
    (activityType: ActivityType) => ({
      System: 'CMS Web',
      Corporation: 'Lai Yih',
      Factory: factory,
      Department: '',
      DocKey:
        activityType.trim() === '4.1'
          ? `${activityType.trim()}${docKey}`
          : activityType.trim(),
      SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
      EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
      ActivityType: activityType.trim(),
      DataType: activityType.trim() === '3.1' ? '1' : '999',
      DocType: '入庫單',
      UndDoc: '',
      DocFlow: '',
      DocDate: docDate,
      DocDate2: docDate,
      DocNo: receivedNo.trim(),
      UndDocNo: '',
      CustVenName: custVenName,
      InvoiceNo: '',
      TransType: transType,
      Departure: departure,
      Destination: destination,
      PortType: portType,
      StPort: portOfDeparture,
      ThPort: '',
      EndPort: portOfArrival,
      Product: matId.trim(),
      Quity: qtyReceive,
      Amount: '',
      ActivityData: activityData,
      ActivityUnit: 'KG',
      Unit: '',
      UnitWeight: '',
      Memo: '',
      CreateDateTime: '',
      Creator: '',
      ActivitySource: matId.charAt(0),
    }),
  );
};
