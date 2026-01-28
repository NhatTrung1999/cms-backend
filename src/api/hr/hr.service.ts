import {
  Inject,
  Injectable,
  InternalServerErrorException,
  NotFoundException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { buildQueryHRModule } from 'src/helper/hrmodule';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs';

@Injectable()
export class HrService {
  constructor(
    @Inject('LYV_HRIS') private readonly LYV_HRIS: Sequelize,
    @Inject('LHG_HRIS') private readonly LHG_HRIS: Sequelize,
    @Inject('LVL_HRIS') private readonly LVL_HRIS: Sequelize,
    @Inject('LYM_HRIS') private readonly LYM_HRIS: Sequelize,
    @Inject('JAZ_HRIS') private readonly JAZ_HRIS: Sequelize,
    @Inject('JZS_HRIS') private readonly JZS_HRIS: Sequelize,
  ) {}

  async findAll(
    dateFrom: string,
    dateTo: string,
    fullName: string,
    id: string,
    department: string,
    joinDate: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'ID',
    sortOrder: string = 'asc',
  ) {
    let db: Sequelize;
    switch (factory.trim().toUpperCase()) {
      case 'LYV':
        db = this.LYV_HRIS;
        break;
      case 'LHG':
        db = this.LHG_HRIS;
        break;
      case 'LVL':
        db = this.LVL_HRIS;
        break;
      case 'LYM':
        db = this.LYM_HRIS;
        break;
      case 'JAZ':
        db = this.JAZ_HRIS;
        break;
      case 'JZS':
        db = this.JZS_HRIS;
        break;
      default:
        throw new Error('Invalid factory code');
    }
    const offset = (page - 1) * limit;
    const { query, countQuery } = buildQueryHRModule(
      dateFrom,
      dateTo,
      fullName,
      id,
      department,
      joinDate,
      factory,
    );
    const replacements: any = [];
    if (dateFrom && dateTo) {
      replacements.push(dateFrom, dateTo);
    }

    if (fullName) {
      replacements.push(`%${fullName}%`);
    }

    if (id) {
      replacements.push(`%${id}%`);
    }

    if (department) {
      replacements.push(
        `${factory.trim().toUpperCase() === 'LYM' ? department : `%${department}%`}`,
      );
    }

    if (joinDate) {
      replacements.push(joinDate);
    }

    const [dataResults, countResults] = await Promise.all([
      db.query(query, { replacements, type: QueryTypes.SELECT }),
      db.query(countQuery, {
        type: QueryTypes.SELECT,
        replacements,
      }),
    ]);

    let data = dataResults;

    data.sort((a, b) => {
      const aValue = a[sortField];
      const bValue = b[sortField];
      if (sortOrder === 'asc') {
        return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
      } else {
        return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
      }
    });

    data = data.slice(offset, offset + limit);

    const total = (countResults[0] as { total: number })?.total || 0;

    const hasMore = offset + data.length < total;
    return { data, page, limit, total, hasMore };
  }

  async findAllDepartment(factory: string) {
    try {
      let db: Sequelize;
      let query: string = '';
      switch (factory.trim().toUpperCase()) {
        case 'LYV':
          db = this.LYV_HRIS;
          query = `SELECT DISTINCT Department_Name  AS label
                          ,Department_Serial_Key     AS value
                    FROM   Data_Department
                    WHERE  NoUse_Date = '2099-12-31 00:00:00.000'
                          AND ISNULL(Hide ,'0')<>'1'
                    ORDER BY Department_Name`;
          break;
        case 'LHG':
          db = this.LHG_HRIS;
          query = `SELECT DISTINCT Department_Name  AS label
                          ,Department_Serial_Key     AS value
                    FROM   Data_Department
                    WHERE  NoUse_Date = '2099-12-31 00:00:00.000'
                          AND Building_Overtime<>''
                    ORDER BY
                          Department_Name`;
          break;
        case 'LVL':
          db = this.LVL_HRIS;
          query = `SELECT DISTINCT Department_Name  AS label
                          ,Department_Serial_Key     AS value
                    FROM   Data_Department
                    WHERE  NoUse_Date = '2099-12-31 00:00:00.000'
                          AND ISNULL(Hide ,'0')<>'1'
                    ORDER BY
                          Department_Name`;
          break;
        case 'LYM':
          db = this.LYM_HRIS;
          query = `SELECT DISTINCT DeptName  AS label
                          ,__sno              AS value
                    FROM   HR_DEPT
                    ORDER BY
                          DeptName`;
          break;
        case 'JAZ':
          db = this.JAZ_HRIS;
          query = `SELECT DISTINCT Department_Name  AS label
                          ,Department_Serial_Key     AS value
                    FROM   Data_Department
                    WHERE  NoUse_Date = '2099-12-31 00:00:00.000'
                    ORDER BY
                          Department_Name`;
          break;
        case 'JZS':
          db = this.JZS_HRIS;
          query = `SELECT DISTINCT Department_Name  AS label
                          ,Department_Serial_Key     AS value
                    FROM   Data_Department
                    WHERE  NoUse_Date = '2099-12-31 00:00:00.000'
                    ORDER BY
                          Department_Name`;
          break;
        default:
          throw new Error('Invalid factory code');
      }

      const results = await db.query(query, { type: QueryTypes.SELECT });
      return results;
    } catch (error) {
      throw new InternalServerErrorException(error?.message);
    }
  }

  async update(
    id: string,
    updateDto: {
      CurrentAddress: string;
      TransportationMethod: string;
    },
    factory: string,
    userid: string,
  ) {
    try {
      let db: Sequelize;
      switch (factory.trim().toUpperCase()) {
        case 'LYV':
          db = this.LYV_HRIS;
          break;
        case 'LHG':
          db = this.LHG_HRIS;
          break;
        case 'LVL':
          db = this.LVL_HRIS;
          break;
        case 'LYM':
          db = this.LYM_HRIS;
          break;
        case 'JAZ':
          db = this.JAZ_HRIS;
          break;
        case 'JZS':
          db = this.JZS_HRIS;
          break;
        default:
          throw new Error('Invalid factory code');
      }

      const query = `
        UPDATE users
        SET
          Address_Live = :address,
          Vehicle = :vehicle,
          UpdatedBy = :updatedBy,
          UpdatedAt = GETDATE()
        OUTPUT 
          INSERTED.userId as id, 
          INSERTED.Address_Live as CurrentAddress, 
          INSERTED.Vehicle as TransportationMethod
        WHERE userId = :id
      `;

      const results = await db.query(query, {
        replacements: {
          address: updateDto.CurrentAddress,
          vehicle: updateDto.TransportationMethod,
          updatedBy: userid,
          id: id,
        },
        type: QueryTypes.SELECT,
      });
      if (!results || results.length === 0) {
        throw new NotFoundException('User not found or ID is incorrect');
      }
      return results[0];
    } catch (error) {
      if (error instanceof NotFoundException) {
        throw error;
      }
      throw new InternalServerErrorException(error.message);
    }
  }

  async importFromExcel(
    file: Express.Multer.File,
    userid: string,
    factory: string,
  ) {
    try {
      let updateCount = 0;
      if (!file.path || !fs.existsSync(file.path)) {
        throw new Error('File path not found!');
      }

      let db: Sequelize;
      switch (factory.trim().toUpperCase()) {
        case 'LYV':
          db = this.LYV_HRIS;
          break;
        case 'LHG':
          db = this.LHG_HRIS;
          break;
        case 'LVL':
          db = this.LVL_HRIS;
          break;
        case 'LYM':
          db = this.LYM_HRIS;
          break;
        case 'JAZ':
          db = this.JAZ_HRIS;
          break;
        case 'JZS':
          db = this.JZS_HRIS;
          break;
        default:
          throw new Error('Invalid factory code');
      }

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(file.path);
      const worksheet = workbook.getWorksheet(1);
      if (!worksheet) {
        throw new Error('No worksheet found in the Excel file');
      }

      const headerRow = worksheet.getRow(1);
      const headers: string[] = [];
      headerRow.eachCell((cell) => {
        if (cell.value) {
          headers.push(cell.value.toString().trim().replace(/\s+/g, '_'));
        }
      });
      const requiredHeaders = [
        'ID',
        'Full_Name',
        'Department',
        'Permanent_Address',
        'Current_Address',
        'Transportation_Method',
      ];

      const missingHeaders = requiredHeaders.filter(
        (h) => !headers.includes(h),
      );

      if (missingHeaders.length > 0) {
        throw new Error(
          `Excel file format is invalid! Missing columns: ${missingHeaders.join(', ')}`,
        );
      }

      const data: {
        ID: string;
        Full_Name: string;
        Department: string;
        Permanent_Address: string;
        Current_Address: string;
        Transportation_Method: string;
      }[] = [];

      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber === 1 && row.cellCount > 0) return;

        const rowData: any = {};
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const headerCell = headerRow.getCell(colNumber);
          const headerValue = headerCell.value;
          if (headerValue !== null && headerValue !== undefined) {
            const key = headerValue.toString().trim().replace(/\s+/g, '_');
            rowData[key] = cell.value;
          }
        });
        if (
          rowData?.ID &&
          rowData?.Full_Name &&
          rowData?.Department &&
          rowData?.Permanent_Address &&
          rowData?.Current_Address &&
          rowData?.Transportation_Method
        ) {
          data.push(rowData);
        }
      });

      const updatedRecords: any[] = [];
      for (let item of data) {
        const records: any[] = await db.query(
          `SELECT userId
          FROM users
          WHERE userId = ?`,
          {
            replacements: [String(item.ID)],
            type: QueryTypes.SELECT,
          },
        );

        if (records && records.length > 0) {
          const queryUpdate = `UPDATE users
                                SET
                                  Address_Live = :address,
                                  Vehicle = :vehicle,
                                  UpdatedBy = :updatedBy,
                                  UpdatedAt = GETDATE()
                                OUTPUT 
                                  INSERTED.userId as id, 
                                  INSERTED.Address_Live as CurrentAddress, 
                                  INSERTED.Vehicle as TransportationMethod
                                WHERE userId = :id`;
          const updateResults: any[] = await db.query(queryUpdate, {
            replacements: {
              address: item.Current_Address,
              vehicle: item.Transportation_Method,
              updatedBy: userid,
              id: String(item.ID),
            },
            type: QueryTypes.SELECT,
          });

          if (updateResults && updateResults.length > 0) {
            updatedRecords.push(updateResults[0]);
          }
        }
        // if (records[0].total > 0) {
        //   await db.query(
        //     `UPDATE users
        //     SET
        //       Address_Live = ?,
        //       Vehicle = ?,
        //       UpdatedBy = ?,
        //       UpdatedAt = GETDATE()
        //     WHERE userId = ?`,
        //     {
        //       replacements: [
        //         item.Current_Address,
        //         item.Transportation_Method,
        //         userid,
        //         String(item.ID),
        //       ],
        //       type: QueryTypes.SELECT,
        //     },
        //   );
        //   updateCount++;
        // }
      }
      // const message = `Updated: ${updateCount} records. Total rows processed: ${data.length}.`;
      return {
        message: `Updated: ${updatedRecords.length} records. Total rows processed: ${data.length}.`,
        updatedData: updatedRecords,
      };
    } catch (error) {
      throw new InternalServerErrorException(`${error.message}`);
    }
  }

  async exportToExcel(
    dateFrom: string,
    dateTo: string,
    fullName: string,
    id: string,
    department: string,
    joinDate: string,
    factory: string,
  ) {
    let db: Sequelize;
    switch (factory.trim().toUpperCase()) {
      case 'LYV':
        db = this.LYV_HRIS;
        break;
      case 'LHG':
        db = this.LHG_HRIS;
        break;
      case 'LVL':
        db = this.LVL_HRIS;
        break;
      case 'LYM':
        db = this.LYM_HRIS;
        break;
      case 'JAZ':
        db = this.JAZ_HRIS;
        break;
      case 'JZS':
        db = this.JZS_HRIS;
        break;
      default:
        throw new Error('Invalid factory code');
    }
    const { query } = await buildQueryHRModule(
      dateFrom,
      dateTo,
      fullName,
      id,
      department,
      joinDate,
      factory,
    );
    const replacements: any = [];
    if (dateFrom && dateTo) {
      replacements.push(dateFrom, dateTo);
    }

    if (fullName) {
      replacements.push(`%${fullName}%`);
    }

    if (id) {
      replacements.push(`%${id}%`);
    }

    if (department) {
      replacements.push(`%${department}%`);
    }

    if (joinDate) {
      replacements.push(joinDate);
    }

    const data = await db.query(query, {
      type: QueryTypes.SELECT,
      replacements,
    });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('List');

    worksheet.columns = [
      {
        header: 'ID',
        key: 'ID',
      },
      {
        header: 'Full Name',
        key: 'FullName',
      },
      {
        header: 'Department',
        key: 'Department',
      },
      {
        header: 'Join Date',
        key: 'JoinDate',
      },
      {
        header: 'Permanent Address',
        key: 'PermanentAddress',
      },
      {
        header: 'Current Address',
        key: 'CurrentAddress',
      },
      {
        header: 'Transportation Method',
        key: 'TransportationMethod',
      },
      {
        header: 'Number of Working Days',
        key: 'Number_of_Working_Days',
      },
    ];

    worksheet.addRows(data);

    worksheet.columns.forEach((column, index) => {
      let maxLength = 10;
      if (typeof column.eachCell === 'function') {
        column.eachCell({ includeEmpty: true }, (cell) => {
          const cellValue = cell.value ? cell.value.toString() : '';
          const cellLength = cellValue.length;

          if (cellLength > maxLength) {
            maxLength = cellLength;
          }
        });
      }
      column.width = Math.min(maxLength + 4, 50);
    });

    return await workbook.xlsx.writeBuffer();
  }
}
