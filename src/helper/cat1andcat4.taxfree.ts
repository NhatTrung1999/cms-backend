import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { memoizeByKey } from './lookup.helper';

interface TaxFreeZoneRow {
  SupplierCode: string;
  TransportationMethod?: string;
  Departure?: string;
}

/**
 * Cat1&4 rows whose supplier has a tax-free zone address (CMW_TAX_FREE_ZONE_ADDRESS, of `factory`) are sent
 * by Land from that address. The address is looked up once per distinct supplier code, not once per row.
 */
export const applyTaxFreeZoneToRows = async <T extends TaxFreeZoneRow>(
  rows: T[],
  factory: string,
  dbEIP: Sequelize | undefined,
): Promise<T[]> => {
  const supplierCodes = rows.map((item) => item.SupplierCode.trim());

  const findTaxFreeZone = memoizeByKey(
    async (supplierCode) =>
      (await dbEIP?.query<{ TaxFreeZoneAddress: string }>(
        `SELECT TaxFreeZoneAddress
          FROM CMW_TAX_FREE_ZONE_ADDRESS
          WHERE SupplierID = ? AND Factory = ?`,
        {
          replacements: [supplierCode, factory],
          type: QueryTypes.SELECT,
        },
      )) ?? [],
  );

  return Promise.all(
    rows.map(async (item, index) => {
      const taxFreeZone = await findTaxFreeZone(supplierCodes[index]);
      return {
        ...item,
        TransportationMethod:
          taxFreeZone.length > 0 ? 'Land' : item.TransportationMethod,
        Departure:
          taxFreeZone.length > 0
            ? taxFreeZone[0].TaxFreeZoneAddress
            : item.Departure,
      };
    }),
  );
};
