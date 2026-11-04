import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { CAT9_AND_CAT12_COLUMNS } from './cat9andcat12.columns';
import {
  buildCat9AndCat12AutoSentQuery,
  buildCat9AndCat12ExportQuery,
  buildCat9AndCat12Filter,
  EXPORT_DATE_COLUMN,
} from './cat9andcat12.queries';
import { fillSheet } from './excel.helper';
import { addPortsToRows } from './shipment.helper';

export const getADataExcelFactoryCat9AndCat12 = async (
  sheet: ExcelJS.Worksheet,
  db: Sequelize,
  dateFrom: string,
  dateTo: string,
  factory: string,
  ry: string,
  dbEIP?: Sequelize,
) => {
  const { where, where1, replacements } = buildCat9AndCat12Filter(
    dateFrom,
    dateTo,
    ry,
    EXPORT_DATE_COLUMN,
  );

  const query = buildCat9AndCat12ExportQuery(where, where1);
  let data: any = await db.query(query, {
    replacements,
    type: QueryTypes.SELECT,
  });

  data = await addPortsToRows(data, dbEIP, factory);

  fillSheet(sheet, CAT9_AND_CAT12_COLUMNS, data);
};

export const buildQueryAutoSentCMS = async (
  dateFrom: string,
  dateTo: string,
  factory: string,
  ry: string,
  db?: Sequelize,
) => {
  const queryAddress = `SELECT ${factory?.trim().toLowerCase() === 'lym'.trim().toLowerCase() ? `N'TSANG YIH Co., Ltd /Polo/ချန်းရိ' as [Address]` : '[Address]'}
                        FROM CMW_Info_Factory
                        WHERE CreatedFactory = '${factory}'`;

  const factoryAddress =
    (await db?.query(queryAddress, {
      type: QueryTypes.SELECT,
    })) || [];

  const { where, where1 } = buildCat9AndCat12Filter(
    dateFrom,
    dateTo,
    ry,
    EXPORT_DATE_COLUMN,
  );
  return buildCat9AndCat12AutoSentQuery(
    factory,
    factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['Address'],
    where,
    where1,
  );
};
