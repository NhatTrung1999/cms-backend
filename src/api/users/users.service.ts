import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { CreateUserDto } from './dto/create-user.dto';
import { UpdateUserDto } from './dto/update-user.dto';

@Injectable()
export class UsersService {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('LYV_ERP') private readonly LYV_ERP: Sequelize,
  ) {}

  async validateUser(userid: string, password?: string, _factory?: string) {
    let query = `
    SELECT *
    FROM CMW_Account
    WHERE UserID = ?
  `;
    const replacements: any[] = [userid];
    if (password) {
      query += ` AND [Password] = ?`;
      replacements.push(password);
    }
    const payload: any = await this.EIP.query<any>(query, {
      replacements,
      type: QueryTypes.SELECT,
    });
    if (payload.length === 0) return false;
    return payload[0];
  }

  async validateErpUser(userid: string, password: string) {
    const payload: any = await this.LYV_ERP.query<any>(
      `
        SELECT USERID, USERNAME, EMAIL, PWD, LOCK
        FROM Busers
        WHERE USERID = ? AND PWD = ?  
      `,
      {
        replacements: [userid, password],
        type: QueryTypes.SELECT,
      },
    );

    if (payload.length === 0) return false;
    return payload[0];
  }

  async checkLock(userid: string) {
    const payload: any = await this.LYV_ERP.query<any>(
      `
      SELECT USERID, USERNAME, EMAIL, PWD, LOCK
      FROM Busers
      WHERE USERID = ? AND LOCK = 'Y'  
    `,
      {
        replacements: [userid],
        type: QueryTypes.SELECT,
      },
    );

    if (payload.length === 1) return false;

    return true;
  }

  async checkExistUser(userid: string) {
    const payload: any = await this.EIP.query<any>(
      `
      SELECT *
      FROM CMW_Account
      WHERE UserID = ?
    `,
      {
        replacements: [userid],
        type: QueryTypes.SELECT,
      },
    );
    if (payload.length === 0) return false;
    return payload[0];
  }

  private async ensureModulePermissionTable() {
    await this.EIP.query(`
      IF OBJECT_ID('dbo.CMW_UserModulePermission', 'U') IS NULL
      BEGIN
        CREATE TABLE dbo.CMW_UserModulePermission (
          ID UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
          UserID NVARCHAR(100) NOT NULL,
          ModulePath NVARCHAR(255) NOT NULL,
          IsAllowed BIT NOT NULL DEFAULT 0,
          UpdatedAt NVARCHAR(100) NULL,
          UpdatedDate DATETIME NOT NULL DEFAULT GETDATE()
        );

        CREATE UNIQUE INDEX UX_CMW_UserModulePermission_User_Module
        ON dbo.CMW_UserModulePermission(UserID, ModulePath);
      END
    `);
  }

  async getUserModulePermissionState(userid: string) {
    await this.ensureModulePermissionTable();

    const rows = await this.EIP.query<any>(
      `
        SELECT ModulePath, IsAllowed
        FROM CMW_UserModulePermission
        WHERE UserID = ?
      `,
      {
        replacements: [userid],
        type: QueryTypes.SELECT,
      },
    );

    return {
      permissionsConfigured: rows.length > 0,
      modulePermissions: rows
        .filter((item) => Boolean(item.IsAllowed))
        .map((item) => item.ModulePath),
    };
  }

  async getModulePermissions(userid: string) {
    return this.getUserModulePermissionState(userid);
  }

  async updateModulePermissions(
    userid: string,
    modulePaths: string[] = [],
    allModulePaths: string[] = [],
    updatedAt?: string,
  ) {
    await this.ensureModulePermissionTable();

    const uniqueAllModulePaths = Array.from(new Set(allModulePaths));
    const allowed = new Set(modulePaths);

    await this.EIP.transaction(async (t) => {
      await this.EIP.query(`DELETE FROM CMW_UserModulePermission WHERE UserID = ?`, {
        replacements: [userid],
        type: QueryTypes.DELETE,
        transaction: t,
      });

      for (const modulePath of uniqueAllModulePaths) {
        await this.EIP.query(
          `
            INSERT INTO CMW_UserModulePermission
              (UserID, ModulePath, IsAllowed, UpdatedAt, UpdatedDate)
            VALUES
              (?, ?, ?, ?, GETDATE())
          `,
          {
            replacements: [
              userid,
              modulePath,
              allowed.has(modulePath) ? 1 : 0,
              updatedAt ?? null,
            ],
            type: QueryTypes.INSERT,
            transaction: t,
          },
        );
      }
    });

    return this.getUserModulePermissionState(userid);
  }

  async getSearch(
    userid: string,
    name: string,
    sortField: string = 'UserID',
    sortOrder: string = 'asc',
  ) {
    let where = `WHERE 1=1`;
    const replacements: any[] = [];

    if (userid) {
      where += ` AND UserID LIKE ?`;
      replacements.push(`%${userid}%`);
    }

    if (name) {
      where += ` AND NAME LIKE ?`;
      replacements.push(`%${name}%`);
    }

    let query = `
      SELECT ID, UserID, Name,Email,Role, Status, CreatedAt, CreatedDate, UpdatedAt, UpdatedDate
      FROM CMW_Account
      ${where}
      ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
    `;

    const result = await this.EIP.query<any>(query, {
      replacements,
      type: QueryTypes.SELECT,
    });

    return result;
  }

  async addUser(createUserDto: CreateUserDto) {
    const { id, userid, name, email, role, status, createdAt } = createUserDto;
    try {
      const result = await this.EIP.query(
        `
        INSERT INTO CMW_Account
        (
          ID,
          Name,
          UserID,
          Email,
          [Role],
          [Status],
          CreatedAt,
          CreatedDate
        )

        OUTPUT INSERTED.ID, INSERTED.Name, INSERTED.UserID, INSERTED.Email,
             INSERTED.Role, INSERTED.Status, INSERTED.CreatedAt, INSERTED.CreatedDate, INSERTED.UpdatedAt, INSERTED.UpdatedDate

        VALUES
        (
          ?,
          ?,
          ?,
          ?,
          ?,
          ?,
          ?,
          GETDATE()
        );
      `,
        {
          replacements: [id, name, userid, email, role, status, createdAt],
          type: QueryTypes.INSERT,
        },
      );
      return result[0];
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  async updateUser(updateUserDto: UpdateUserDto) {
    const { name, email, role, status, updatedAt, id } = updateUserDto;
    try {
      await this.EIP.query(
        `
        UPDATE CMW_Account
        SET 
            Name = ?,
            Email = ?,
            [Role] = ?,
            [Status] = ?,
            UpdatedAt = ?,
            UpdatedDate = GETDATE()
        WHERE ID = ?
      `,
        {
          replacements: [name, email, role, status, updatedAt, id],
          type: QueryTypes.UPDATE,
        },
      );

      const payload: any = await this.EIP.query(`SELECT * FROM CMW_Account`, {
        type: QueryTypes.SELECT,
      });
      return payload;
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  async deleteUser(id: string) {
    try {
      await this.EIP.query(`DELETE FROM CMW_Account WHERE ID = ?`, {
        replacements: [id],
        type: QueryTypes.DELETE,
      });

      const payload: any = await this.EIP.query(`SELECT * FROM CMW_Account`, {
        type: QueryTypes.SELECT,
      });
      return payload;
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }
}
