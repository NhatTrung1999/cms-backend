import dayjs from 'dayjs';
import {
  ACTIVITY_TYPES as ACTIVITY_TYPES_CAT7,
  ActivityType as ActivityTypeCat7,
} from 'src/types/cat7';

/** A row of the Cat7 (commuting) CMS query (only the columns the mapper reads). */
export interface Cat7CmsRow {
  Staff_ID?: string;
  Main_transportation_type?: string;
  Number_of_working_days?: number | string;
  Factory_address?: string;
  Residential_address?: string;
  Factory_Name?: string;
  Department_Name?: string;
}

// Each mapper turns one row read from a factory database into the CMS payload row(s).
// It returns [] when `dockeyCMS` is not one of the activity types of its category.

export const mapToCMSFormatCat7 = (
  item: Cat7CmsRow,
  dateFrom: string,
  dateTo: string,
  dockeyCMS: string,
) => {
  const staffId = item.Staff_ID ?? '';
  const Main_transportation_type = item.Main_transportation_type ?? '';
  const Number_of_working_days = item.Number_of_working_days ?? '';
  const Factory_address = item.Factory_address ?? '';
  const Residential_address = item.Residential_address ?? Factory_address;
  const FactoryName = item.Factory_Name ?? '';
  const DepartmentName = item.Department_Name ?? '';

  let dockey = '';
  let transType = '';

  switch (Main_transportation_type.trim().toLowerCase()) {
    case 'Walking'.trim().toLowerCase(): //步行
      dockey = '3.3.1';
      transType = '步行';
      break;
    case 'Bicycle'.trim().toLowerCase(): //自行車
      dockey = '3.3.2';
      transType = '自行車';
      break;
    case 'Electric motorcycle'.trim().toLowerCase(): //電動機車
      dockey = '3.3.3';
      transType = '電動機車';
      break;
    case 'Motorcycle'.trim().toLowerCase(): //機車
      dockey = '3.3.4';
      transType = '機車';
      break;
    case 'Electric car'.trim().toLowerCase(): //電動汽車
      dockey = '3.3.5';
      transType = '電動汽車';
      break;
    case 'Car'.trim().toLowerCase(): //汽車
      dockey = '3.3.6';
      transType = '汽車';
      break;
    case 'Bus'.trim().toLowerCase(): //公車/客運
      dockey = '3.3.7';
      transType = '公車/客運';
      break;
    case 'Company shuttle bus'.trim().toLowerCase(): //員工接駁車
      dockey = '3.3.8';
      transType = '員工接駁車';
      break;
    case 'Subway'.trim().toLowerCase(): //地鐵
      dockey = '3.3.9';
      transType = '地鐵';
      break;
    default:
      break;
  }

  return ACTIVITY_TYPES_CAT7.filter((item) => item === dockeyCMS).map(
    (activityType: ActivityTypeCat7) => ({
      System: 'CMS Web',
      Corporation: 'LAI YIH',
      Factory: FactoryName,
      Department: DepartmentName.slice(0, 20),
      DocKey: dockey,
      SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
      EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
      ActivityType: activityType,
      DataType: '2',
      DocType: 'CMS Web',
      DocDate: dayjs('2025/12/31').format('YYYY/MM/DD'),
      DocDate2: dayjs('2025/12/31').format('YYYY/MM/DD'),
      UndDocNo: staffId,
      TransType: transType,
      Departure: Residential_address,
      Destination: Factory_address,
      Attendance: Number_of_working_days.toString(),
      Memo: '',
      CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
      Creator: '',
    }),
  );
};
