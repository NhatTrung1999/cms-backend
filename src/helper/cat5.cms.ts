import dayjs from 'dayjs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { ACTIVITY_TYPES, ActivityType } from 'src/types/cat5';
import { buildQueryAutoSentCMS } from './cat5.helper';

/** A row of the Cat5 (waste) CMS query (only the columns the mapper reads). */
export interface Cat5CmsRow {
  Vendor_ID?: string;
  Factory_address?: string;
  Waste_collection_address?: string;
  The_type_of_waste?: string;
  Waste_type?: string;
  Weight_of_waste_treated_Unit_kg?: number;
  Waste_Treatment_method: string;
  Factory_Name?: string;
  Waste_disposal_date?: string | Date;
}

/**
 * Turns one Cat5 (waste) row read from a factory database into the CMS payload row(s).
 * It returns [] when `dockeyCMS` is not one of the activity types of the category.
 */
export const mapToCMSFormatCat5 = (
  item: Cat5CmsRow,
  dateFrom: string,
  dateTo: string,
  dockeyCMS: string,
) => {
  const custVenName = item.Vendor_ID;
  const departure = item.Factory_address;
  const destination = item.Waste_collection_address;
  const product = `${item.The_type_of_waste} - ${item.Waste_type}`;
  const activityData = item.Weight_of_waste_treated_Unit_kg;
  const wasteTreatmentMethod = item.Waste_Treatment_method;
  const factoryName = item.Factory_Name;
  const wasteDisposalDate = item.Waste_disposal_date;
  let dockey = '';

  switch (wasteTreatmentMethod.trim().toLowerCase()) {
    case 'Waste Recycling'.trim().toLowerCase():
      dockey = '4.4.01';
      break;
    case 'Waste to Energy'.trim().toLowerCase():
      dockey = '4.4.02';
      break;
    case 'Landfill'.trim().toLowerCase():
      dockey = '4.4.03';
      break;
    case 'Waste Incineration'.trim().toLowerCase():
      dockey = '4.4.04';
      break;
    case 'Waste other treatment method'.trim().toLowerCase():
      dockey = '4.4.05';
      break;
    default:
      dockey = '';
      break;
  }

  return ACTIVITY_TYPES.filter((item) => item === dockeyCMS).map(
    (activityType: ActivityType) => ({
      System: 'CMW', //default
      Corporation: 'LAI YIH', //default
      Factory: factoryName,
      Department: '',
      DocKey: activityType.trim() === '3.6' ? '3.6' : dockey,
      SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
      EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
      ActivityType: activityType, //default
      DataType: activityType.trim() === '3.6' ? '1' : '999', //default
      DocType: 'CMS Web', //default
      UndDoc: '',
      DocFlow: '',
      DocDate: dayjs(wasteDisposalDate).format('YYYY/MM/DD'),
      DocDate2: dayjs(wasteDisposalDate).format('YYYY/MM/DD'),
      DocNo: '',
      UndDocNo: '',
      CustVenName: custVenName,
      InvoiceNo: '',
      TransType: '陸運', //default
      Departure: departure,
      Destination: destination,
      PortType: '',
      StPort: '',
      ThPort: '',
      EndPort: '',
      Product: product,
      Quity: '',
      Amount: '',
      ActivityData: activityData,
      ActivityUnit: 'KG', //default
      Unit: '',
      UnitWeight: '',
      Memo: '',
      CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
      Creator: '',
      AccNo: '6002',
      AccName: '其他費用',
      ActivitySource: wasteTreatmentMethod,
    }),
  );
};

/**
 * CMS payload rows of one factory: reads its rows of the date range (WMS `db`) and maps them with mapToCMSFormatCat5.
 */
export const getCat5CMSRows = async (
  db: Sequelize,
  dbEIP: Sequelize,
  factory: string,
  dateFrom: string,
  dateTo: string,
  dockeyCMS: string,
) => {
  const query = await buildQueryAutoSentCMS(dateFrom, dateTo, factory, dbEIP);

  const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

  const data = await db.query<any>(query, {
    type: QueryTypes.SELECT,
    replacements,
  });

  return data.flatMap((item) =>
    mapToCMSFormatCat5(item, dateFrom, dateTo, dockeyCMS),
  );
};
