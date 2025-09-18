import {
  Inject,
  Injectable,
  UnauthorizedException,
  InternalServerErrorException,
} from '@nestjs/common';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';

@Injectable()
export class Cat9andcat12Service {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('ERP') private readonly ERP: Sequelize,
  ) {}

  // async getData(date, userID) {

  //   console.log(date, userID);

  //   if (!userID) throw new UnauthorizedException();

  //   try {
  //     let where = ' 1 = 1 ';

  //     if (date !== '') {
  //     }

  //     const payload: any = await this.EIP.query<any>(
  //       `
  //         SELECT *
  //         FROM CMS_File_Management
  //         WHERE CreatedAt = '${userID}' ${where}
  //       `,
  //       {
  //         replacements: [],
  //         type: QueryTypes.SELECT,
  //       },
  //     );
  //     return payload;
  //   } catch (error: any) {
  //     throw new InternalServerErrorException(error);
  //   }
  // }

  async getData(date: string, offset: number = 1, limit: number = 20) {
    try {
      let query = `
        SELECT im.INV_DATE,im.INV_NO,id.STYLE_NAME,p.Qty,p.GW,im.CUSTID,
            'Truck' LocalLandTransportation,sb.Place_Delivery,im.TO_WHERE Country
            ,ISNULL(ISNULL(bg.SHPIDS,CASE WHEN (do.ShipMode='Air') AND (do.Shipmode_1 IS NULL) THEN '10 AC'
                          WHEN (do.ShipMode='Air Expres') AND (do.Shipmode_1 IS NULL) THEN '20 CC'
                                            WHEN (do.ShipMode='Ocean') AND (do.Shipmode_1 IS NULL) THEN '11 SC'
                                            WHEN do.ShipMode_1 IS NULL THEN ''
                                            ELSE do.ShipMode_1 END),y.ShipMode) TransportMethod
              
        FROM INVOICE_M im
        LEFT JOIN Ship_Booking sb ON sb.INV_NO = im.INV_NO
        LEFT JOIN INVOICE_D AS id ON id.INV_NO = im.INV_NO
        LEFT JOIN (SELECT INV_NO,RYNO,SUM(PAIRS)Qty,SUM (GW) GW
                  FROM PACKING
                  GROUP BY INV_NO,RYNO) p ON p.INV_NO=id.INV_NO AND p.RYNO=id.RYNO
        LEFT JOIN YWDD y ON y.DDBH=id.RYNO
        LEFT JOIN DE_ORDERM do ON do.ORDERNO=y.YSBH
        LEFT JOIN B_GradeOrder bg ON bg.ORDER_B=y.YSBH
        WHERE 1=1 ${date ? `AND CONVERT(VARCHAR, im.INV_DATE, 23) = '${date}'` : ''}
        ORDER BY im.INV_DATE 
        OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY
      `;
      // const replacements: any[] = [];

      // if (date) {
      //   query += ` AND CONVERT(VARCHAR, im.INV_DATE, 23) = ?`;
      //   // replacements.push(date);
      // }

      // query += ` ORDER BY im.INV_DATE
      //           OFFSET ? ROWS FETCH NEXT ? ROWS ONLY`;
      // replacements.push(offset, limit)

      const res = await this.ERP.query(query, {
        replacements: [],
        type: QueryTypes.SELECT,
      });

      return res;
    } catch (error) {
      throw new InternalServerErrorException(error);
    }
  }
}
