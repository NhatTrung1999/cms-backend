import { Inject, Injectable } from '@nestjs/common';
import { CreateCat6Dto } from './dto/create-cat6.dto';
import { UpdateCat6Dto } from './dto/update-cat6.dto';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';
import { ICat6Data } from 'src/types/cat6';

@Injectable()
export class Cat6Service {
  constructor(@Inject('UOF') private readonly UOF: Sequelize) {}

  async getDataCat6() {
    const records: ICat6Data[] = await this.UOF.query(
      `SELECT *
        FROM CDS_HRBUSS_BusTripData
        WHERE DOC_NBR = 'LYV-HR-BT251000015'`,
      { type: QueryTypes.SELECT },
    );
    return records;
  }
}
