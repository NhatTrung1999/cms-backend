import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { autoFitColumnsAndAddBorders } from './excel.helper';
import { getDayAfter } from './date.helper';
import { applyTaxFreeZoneToRows } from './cat1andcat4.taxfree';
import { memoizeByKey } from './lookup.helper';
import {
  buildCat1AndCat4Base,
  getSupplierExclusion,
} from './cat1andcat4.queries';

const factoryAddressCache = new Map<string, string>();
const styleAutoFillCache = new Map<string, string | null>();

const getFactoryAddress = async (factory: string, db?: Sequelize) => {
  const normalizedFactory = String(factory ?? '')
    .trim()
    .toUpperCase();
  if (!normalizedFactory) {
    return 'N/A';
  }

  const cachedAddress = factoryAddressCache.get(normalizedFactory);
  if (cachedAddress) {
    return cachedAddress;
  }

  const rows =
    (await db?.query(
      `SELECT TOP 1 [Address]
       FROM CMW_Info_Factory
       WHERE CreatedFactory = :factory`,
      {
        replacements: { factory: normalizedFactory },
        type: QueryTypes.SELECT,
      },
    )) || [];

  const address = rows.length > 0 ? String(rows[0]['Address'] ?? 'N/A') : 'N/A';
  factoryAddressCache.set(normalizedFactory, address);
  return address;
};

const getStyleAutoFillByPrefix = async (prefix: string, db?: Sequelize) => {
  const normalizedPrefix = String(prefix ?? '')
    .trim()
    .toUpperCase();
  if (!normalizedPrefix) {
    return null;
  }

  const cachedStyle = styleAutoFillCache.get(normalizedPrefix);
  if (styleAutoFillCache.has(normalizedPrefix)) {
    return cachedStyle ?? null;
  }

  const rows =
    (await db?.query(
      `SELECT TOP 1 Style
       FROM CMW_StyleAutoFill
       WHERE PrefixOfMatCode = :prefix`,
      {
        replacements: { prefix: normalizedPrefix },
        type: QueryTypes.SELECT,
      },
    )) || [];

  const style = rows.length > 0 ? String(rows[0]['Style'] ?? '') : null;
  styleAutoFillCache.set(normalizedPrefix, style);
  return style;
};

export const buildQuery = async (
  sortField: string = 'No',
  sortOrder: string = 'asc',
  factory: string,
  usage: boolean = false,
  unitWeight: boolean = false,
  weight: boolean = false,
  departure: boolean = false,
  db?: Sequelize,
  isExport: boolean = false,
) => {
  const factoryAddress = await getFactoryAddress(factory, db);

  const pagingSql = isExport
    ? ''
    : `OFFSET :offset ROWS
      FETCH NEXT :limit ROWS ONLY;`;

  let where = `WHERE 1=1`;
  if (usage) {
    where += ` AND ISNULL(QtyUsage, 0) = 0`;
  }

  if (unitWeight) {
    where += ` AND ISNULL(UnitWeight, 0) = 0`;
  }

  if (weight) {
    where += ` AND ISNULL(WeightUnitkg, 0) = 0`;
  }

  if (departure) {
    where += ` AND ISNULL(Departure, '') = ''`;
  }

  const notIn = getSupplierExclusion(factory);

  const query = `${buildCat1AndCat4Base(factory, factoryAddress, notIn)}
            ${where}
            ORDER BY
                ${sortField === 'No' ? '[No]' : sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
            ${pagingSql}`;
  return query;
};

export const getADataExcelFactoryCat1AndCat4 = async (
  sheet: ExcelJS.Worksheet,
  db: Sequelize,
  dateFrom: string,
  dateTo: string,
  factory: string,
  usage?: boolean,
  unitWeight?: boolean,
  weight?: boolean,
  departure?: boolean,
  dbEIP?: Sequelize,
) => {
  const query = await buildQuery(
    'No',
    'asc',
    factory,
    usage,
    unitWeight,
    weight,
    departure,
    dbEIP,
    true,
  );
  const replacements = {
    startDate: dateFrom,
    endDate: getDayAfter(dateTo),
  };

  let data: any[] = await db.query(query, {
    replacements,
    type: QueryTypes.SELECT,
  });

  data = await applyTaxFreeZoneToRows(data, factory, dbEIP);

  const styleOfPrefix = memoizeByKey((prefix) =>
    getStyleAutoFillByPrefix(prefix, dbEIP),
  );
  data = await Promise.all(
    data.map(async (item: any) => {
      const prefix = String(item?.MatID ?? '')
        .trim()
        .charAt(0);
      const styleAutoFill = await styleOfPrefix(prefix);
      return {
        ...item,
        Style:
          (item.Style === '' ? styleAutoFill : item?.Style) ||
          (item?.Style ?? styleAutoFill ?? 'N/A'),
      };
    }),
  );

  sheet.columns = [
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

  data.forEach((item) => sheet.addRow(item));

  autoFitColumnsAndAddBorders(sheet);
};

export const buildQueryAutoSentCMS = async (
  factory: string,
  db?: Sequelize,
) => {
  const factoryAddress =
    factory?.trim().toLowerCase() === 'lym'.trim().toLowerCase()
      ? 'TSANG YIH Co., Ltd /Polo/ချန်းရိ'
      : await getFactoryAddress(factory, db);

  const notIn = getSupplierExclusion(factory);

  const query = buildCat1AndCat4Base(factory, factoryAddress, notIn);
  return query;
};
