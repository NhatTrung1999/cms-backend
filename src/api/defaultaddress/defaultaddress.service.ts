import {
  Inject,
  Injectable,
  InternalServerErrorException,
  NotFoundException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { orderByClause } from 'src/helper/sql.helper';
import { Sequelize } from 'sequelize-typescript';
import { v4 as uuidv4 } from 'uuid';
import { readImportRows } from 'src/helper/excel-import.helper';
import { getOwn } from 'src/helper/object.helper';

@Injectable()
export class DefaultaddressService {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('LYV_HRIS') private readonly LYV_HRIS: Sequelize,
    @Inject('LHG_HRIS') private readonly LHG_HRIS: Sequelize,
    @Inject('LVL_HRIS') private readonly LVL_HRIS: Sequelize,
    @Inject('LYM_HRIS') private readonly LYM_HRIS: Sequelize,
    @Inject('JAZ_HRIS') private readonly JAZ_HRIS: Sequelize,
    @Inject('JZS_HRIS') private readonly JZS_HRIS: Sequelize,
  ) {}

  private get factoryDbMap(): Record<string, Sequelize> {
    return {
      LYV: this.LYV_HRIS,
      LHG: this.LHG_HRIS,
      LVL: this.LVL_HRIS,
      LYM: this.LYM_HRIS,
      JAZ: this.JAZ_HRIS,
      JZS: this.JZS_HRIS,
    };
  }

  async getDefaultAddress(sortField: string = 'No', sortOrder: string = 'asc') {
    const records: any[] = await this.EIP.query(
      `SELECT ROW_NUMBER() OVER (ORDER BY ID) AS No,*
        FROM CMW_Default_Address
        ORDER BY ${orderByClause(sortField, sortOrder)}
        `,
      { type: QueryTypes.SELECT },
    );
    return records;
  }

  async updateDefaultAddress(
    id: string,
    updateDto: {
      DefaultAddress: string;
    },
    factory: string,
    userid: string,
  ) {
    try {
      const query = `
              UPDATE CMW_Default_Address
              SET
                DefaultAddress = :defaultAddress,
                UpdatedBy = :updatedBy,
                UpdatedFactory = :updatedFactory,
                UpdatedAt = GETDATE()
              OUTPUT
                INSERTED.ID AS ID,
                INSERTED.DefaultAddress AS DefaultAddress,
                INSERTED.UpdatedBy AS UpdatedBy,
                INSERTED.UpdatedAt AS UpdatedAt
              WHERE ID = :id
            `;

      const results = await this.EIP.query(query, {
        replacements: {
          defaultAddress: updateDto.DefaultAddress,
          updatedBy: userid,
          updatedFactory: factory,
          id: id,
        },
        type: QueryTypes.SELECT,
      });
      if (!results || results.length === 0) {
        throw new NotFoundException('Update failed!');
      }
      return results[0];
    } catch (error) {
      if (error instanceof NotFoundException) {
        throw error;
      }
      throw new InternalServerErrorException(error.message);
    }
  }

  async deleteDefaultAddress(id: string) {
    try {
      await this.EIP.query(`DELETE FROM CMW_Default_Address WHERE ID = ?`, {
        replacements: [id],
        type: QueryTypes.DELETE,
      });
      return { message: 'Deleted successfully!', ID: id };
    } catch (error) {
      throw new InternalServerErrorException(error.message);
    }
  }

  async importExcelDefaultAddress(
    file: Express.Multer.File,
    userid: string,
    factory: string,
  ) {
    try {
      let insertCount = 0;
      const data = await readImportRows<{
        Factory: string;
        Default_Address: string;
      }>(
        file,
        ['Factory', 'Default_Address'],
        (rowData) => rowData?.Factory || rowData?.Default_Address,
      );

      for (const item of data) {
        const id = uuidv4();
        await this.EIP.query(
          `INSERT INTO CMW_Default_Address
              (
                ID,
                Factory,
                DefaultAddress,
                CreatedBy,
                CreatedAt,
                CreatedFactory
              )
              VALUES
              (
                ?,
                ?,
                ?,
                ?,
                GETDATE(),
                ?
              )`,
          {
            replacements: [
              id,
              item.Factory,
              item.Default_Address,
              userid,
              factory,
            ],
            type: QueryTypes.INSERT,
          },
        );
        insertCount++;
      }
      const records: any = await this.EIP.query(
        `SELECT ROW_NUMBER() OVER (ORDER BY ID) AS No, *
          FROM CMW_Default_Address`,
        { type: QueryTypes.SELECT },
      );
      const message = `Processed successfully! Inserted: ${insertCount} records.`;
      return { message, records };
    } catch (error: any) {
      throw new InternalServerErrorException(`${error.message}`);
    }
  }

  async syncDefaultAddress(
    userid: string,
    factory: string,
    syncDefaultAddress: string,
  ) {
    try {
      const db = getOwn(this.factoryDbMap, factory.trim().toUpperCase());
      if (!db) throw new Error('Invalid factory code');
      await db.query(
        `UPDATE users
          SET    Address_Live     = :address
                ,UpdatedBy        = :updatedBy
                ,UpdatedAt        = GETDATE()
          WHERE  Address_Live IS NULL
                AND Vehicle<>'Company shuttle bus'`,
        {
          replacements: {
            address: syncDefaultAddress,
            updatedBy: userid,
          },
        },
      );
      return { message: 'Sync default address successfully!' };
    } catch (error) {
      throw new InternalServerErrorException(`${error.message}`);
    }
  }
}
