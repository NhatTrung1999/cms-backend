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
    @Inject('ERP') private readonly ERP: Sequelize,
  ) {}

  async validateUser(userid: string, password?: string, factory?: string) {
    // console.log(password);
    let query = `
    SELECT *
    FROM CMS_Account
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
    // const payload: any = await this.EIP.query<any>(
    // `
    //   SELECT *
    //   FROM CMS_Account
    //   WHERE UserID = ? AND [Password] = ?
    // `,
    //   {
    //     replacements: [userid, password],
    //     type: QueryTypes.SELECT,
    //   },
    // );
    if (payload.length === 0) return false;
    return payload[0];
  }

  async validateErpUser(userid: string, password: string) {
    const payload: any = await this.ERP.query<any>(
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
    const payload: any = await this.ERP.query<any>(
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
      FROM CMS_Account
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

  async getAll() {
    const result = await this.EIP.query<any>(
      `SELECT UserID, Name,Email,Role, Status, CreatedAt, CreatedDate, UpdatedAt, UpdatedDate
        FROM CMS_Account`,
      { type: QueryTypes.SELECT },
    );

    return result;
  }

  async getSearch(userid: string, name: string) {
    let query = `
      SELECT UserID, Name,Email,Role, Status, CreatedAt, CreatedDate, UpdatedAt, UpdatedDate
      FROM CMS_Account
      WHERE 1=1
    `;
    const replacements: any[] = [];
    if (userid) {
      query += ` AND UserID LIKE ?`;
      replacements.push(`%${userid}%`);
    }

    if (name) {
      query += ` AND NAME LIKE ?`;
      replacements.push(`%${name}%`);
    }

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
        INSERT INTO CMS_Account
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
        UPDATE CMS_Account
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
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  async deleteUser(id: string) {
    try {
      await this.EIP.query(`DELETE FROM CMS_Account WHERE ID = ?`, {
        replacements: [id],
        type: QueryTypes.DELETE,
      });
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }
}
