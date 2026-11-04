import { QueryTypes, Transaction } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { v4 as uuidv4 } from 'uuid';
import { getFactoryCodeByLabel } from 'src/helper/factory.helper';

/**
 * One column of a log table INSERT: the SQL column, the name of its named parameter and where its value comes from.
 * `source` is a field of the sent row, or one of: '@id' (a new uuid), '@fac' (factory code of the row's Factory label),
 * '@user' (the user that sends the log), '@factory' (the factory of that user).
 * A column with `sql` has no parameter (its value is that SQL expression).
 */
type InsertColumn =
  | { column: string; param: string; source: string }
  | { column: string; sql: string };

const lowerFirst = (name: string) =>
  name.charAt(0).toLowerCase() + name.slice(1);

/** A column whose value is `field` of the row (parameter and column named after the field unless given). */
const field = (
  name: string,
  param = lowerFirst(name),
  column = name,
): InsertColumn => ({
  column,
  param,
  source: name,
});
const id: InsertColumn = { column: 'ID', param: 'id', source: '@id' };
const fac: InsertColumn = { column: 'Fac', param: 'fac', source: '@fac' };
const createdUser: InsertColumn = {
  column: 'CreatedUser',
  param: 'createdUser',
  source: '@user',
};
const createdFactory: InsertColumn = {
  column: 'CreatedFactory',
  param: 'createdFactory',
  source: '@factory',
};
const createdAt: InsertColumn = { column: 'CreatedAt', sql: 'GETDATE()' };
const system = field('System', undefined, '[System]');

// Cat1&4 and Cat5 share their table layout.
const CAT1_4_AND_5_COLUMNS: InsertColumn[] = [
  id,
  system,
  field('Corporation'),
  field('Factory'),
  fac,
  field('Department'),
  field('DocKey'),
  field('SPeriodData', 'sPeriod'),
  field('EPeriodData', 'ePeriod'),
  field('ActivityType'),
  field('DataType'),
  field('DocType'),
  field('UndDoc'),
  field('DocFlow'),
  field('DocDate'),
  field('DocDate2'),
  field('DocNo'),
  field('UndDocNo'),
  field('CustVenName'),
  field('InvoiceNo'),
  field('TransType'),
  field('Departure'),
  field('Destination'),
  field('PortType'),
  field('StPort'),
  field('ThPort'),
  field('EndPort'),
  field('Product'),
  field('Quity'),
  field('Amount'),
  field('ActivityData'),
  field('ActivityUnit'),
  field('Unit'),
  field('UnitWeight'),
  field('Memo'),
  field('CreateDateTime'),
  field('Creator'),
  createdUser,
  createdFactory,
  createdAt,
  field('ActivitySource'),
];

const CAT6_COMMON_HEAD: InsertColumn[] = [
  id,
  system,
  field('Corporation'),
  field('Factory'),
  fac,
  field('Department'),
  field('DocKey'),
  field('ActivitySource'),
  field('SPeriodData'),
  field('EPeriodData'),
  field('ActivityType'),
  field('DataType'),
];
const CAT6_COMMON_TAIL: InsertColumn[] = [
  field('Memo'),
  field('CreateDateTime'),
  field('Creator'),
  createdUser,
  createdFactory,
  createdAt,
];

export interface LogInsert {
  table: string;
  columns: InsertColumn[];
  /** Rows per INSERT statement; without it every row has its own statement. */
  chunkSize?: number;
}

export const LOG_INSERT: Record<
  | 'cat1And4'
  | 'cat5'
  | 'cat6BusinessTravel'
  | 'cat6Accommodation'
  | 'cat7'
  | 'cat9And12',
  LogInsert
> = {
  cat1And4: {
    table: 'CMW_Category_1_And_4_Log',
    columns: CAT1_4_AND_5_COLUMNS,
    chunkSize: 40,
  },
  cat5: { table: 'CMW_Category_5_Log', columns: CAT1_4_AND_5_COLUMNS },
  cat6BusinessTravel: {
    table: 'CMW_Category_6_Business_Travel',
    columns: [
      ...CAT6_COMMON_HEAD,
      field('DocType'),
      field('DocDate'),
      field('DocDate2'),
      field('DocNo'),
      field('UndDocNo'),
      field('TransType'),
      field('Departure'),
      field('Destination'),
      ...CAT6_COMMON_TAIL,
    ],
  },
  cat6Accommodation: {
    table: 'CMW_Category_6_Accommodation',
    columns: [
      ...CAT6_COMMON_HEAD,
      field('DocType'),
      field('DocDate'),
      field('DocDate2'),
      field('DocNo'),
      field('UndDocNo'),
      field('TransType'),
      field('ActivityData'),
      ...CAT6_COMMON_TAIL,
    ],
  },
  cat7: {
    table: 'CMW_Category_7_Log',
    columns: [
      id,
      system,
      field('Corporation'),
      field('Factory'),
      fac,
      field('Department'),
      field('DocKey'),
      field('SPeriodData'),
      field('EPeriodData'),
      field('ActivityType'),
      field('DataType'),
      field('DocType'),
      field('DocDate'),
      field('DocDate2'),
      field('UndDocNo'),
      field('TransType'),
      field('Departure'),
      field('Destination'),
      field('Attendance'),
      field('Memo'),
      field('CreateDateTime'),
      field('Creator'),
      createdUser,
      createdFactory,
      createdAt,
    ],
  },
  cat9And12: {
    table: 'CMW_Category_9_And_12_Log',
    columns: [
      id,
      system,
      field('Corporation'),
      field('Factory'),
      fac,
      field('Department'),
      field('DocKey'),
      field('SPeriodData'),
      field('EPeriodData'),
      field('ActivityType'),
      field('DataType'),
      field('DocType'),
      field('UndDoc'),
      field('DocFlow'),
      field('DocDate'),
      field('DocDate2'),
      field('DocNo'),
      field('UndDocNo'),
      field('CustVenName'),
      field('InvoiceNo'),
      field('TransType'),
      field('Departure'),
      field('Destination'),
      field('PortType'),
      field('StPort'),
      field('ThPort'),
      field('EndPort'),
      field('Product'),
      field('Quity'),
      field('Amount'),
      field('ActivityData'),
      field('ActivityUnit'),
      field('Unit'),
      field('UnitWeight'),
      field('Memo'),
      field('CreateDateTime'),
      field('Creator'),
      field('ActivitySource'),
      createdUser,
      createdFactory,
      createdAt,
    ],
  },
};

/**
 * Inserts the rows sent by the page into a log table, inside `transaction`:
 * one statement per row, or one statement per `chunkSize` rows (the parameters of row n end with n).
 */
export const insertLogRows = async (
  eip: Sequelize,
  transaction: Transaction,
  { table, columns, chunkSize }: LogInsert,
  rows: object[],
  sender: { factory: string; userid: string },
) => {
  const valueOf = (row: Record<string, unknown>, source: string) => {
    switch (source) {
      case '@id':
        return uuidv4();
      case '@fac':
        return getFactoryCodeByLabel(row.Factory as string);
      case '@user':
        return sender.userid;
      case '@factory':
        return sender.factory;
      default:
        return row[source];
    }
  };

  const valuesSql = (suffix: string) =>
    `(
      ${columns.map((c) => ('sql' in c ? c.sql : `:${c.param}${suffix}`)).join(',\n      ')}
    )`;

  const insertSql = (values: string, end: string) => `
    INSERT INTO ${table}
    (
      ${columns.map((c) => c.column).join(',\n      ')}
    )
    VALUES
    ${values}${end}
    `;

  const rowsPerStatement = chunkSize ?? 1;
  for (let start = 0; start < rows.length; start += rowsPerStatement) {
    const chunk = rows.slice(start, start + rowsPerStatement) as Record<
      string,
      unknown
    >[];
    const replacements: Record<string, unknown> = {};
    chunk.forEach((row, index) => {
      const suffix = chunkSize ? `${start + index}` : '';
      for (const c of columns) {
        if ('param' in c)
          replacements[`${c.param}${suffix}`] = valueOf(row, c.source);
      }
    });

    const values = chunkSize
      ? chunk.map((_, index) => valuesSql(`${start + index}`)).join(',\n')
      : valuesSql('');
    await eip.query(insertSql(values, chunkSize ? ';' : ''), {
      type: QueryTypes.INSERT,
      transaction,
      replacements: replacements as Record<string, string | number>,
    });
  }
};
