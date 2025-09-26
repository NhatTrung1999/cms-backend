import { Inject, Injectable } from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

@Injectable()
export class Cat5Service {
  constructor(
    @Inject('LYV_WMS') private readonly LYV_WMS: Sequelize,
    @Inject('LHG_WMS') private readonly LHG_WMS: Sequelize,
    @Inject('LYM_WMS') private readonly LYM_WMS: Sequelize,
    @Inject('LVL_WMS') private readonly LVL_WMS: Sequelize,
    @Inject('JAZ_WMS') private readonly JAZ_WMS: Sequelize,
    @Inject('JZS_WMS') private readonly JZS_WMS: Sequelize,
  ) {}

  private readonly querySQL = `SELECT '' DateAsString
                                    ,dwo.WASTE_DATE
                                    ,dwo.VEHICLE_ID
                                    ,dwo.WASTE_CODE
                                    ,dwc.WASTE_NAME_LOCAL+'/'+dwc.WASTE_NAME_ENGLISH WASTE_NAME
                                    ,dwo.QUANTITY
                                    ,dwo.TREATMENT_SUPPLIER                SUPPLIER_ID
                                    ,dtv.TREATMENT_VENDOR_NAME             SUPPLIER
                                    ,dwo.TREATMENT_METHOD_ID               METHOD_ID
                                    ,dtm.TREATMENT_METHOD_ENGLISH_NAME     METHOD
                                    ,dwo.EMERET_WASTE_ID                   EMERET_ID
                                    ,dewc.EMERET_WASTE_ENGLISH_NAME        EMERET
                                    ,dwo.LOCATION_CODE
                                    ,td.ADDRESS+'('+CONVERT(VARCHAR(5) ,td.DISTANCE)+'km)' ADDRESS
                                    ,dwo.HAZARDOUS
                                    ,dwo.NON_HAZARDOUS
                                    ,dwo.LOCKED
                              FROM   dbo.DATA_WASTE_OUTPUT_CUSTOMER dwo
                                    LEFT JOIN dbo.DATA_TREATMENT_VENDOR dtv
                                          ON  dtv.TREATMENT_VENDOR_ID = dwo.TREATMENT_SUPPLIER
                                    LEFT JOIN dbo.DATA_TREATMENT_METHOD dtm
                                          ON  dtm.TREATMENT_METHOD_ID = dwo.TREATMENT_METHOD_ID
                                    LEFT JOIN dbo.DATA_EMERET_WASTE_CATEGORY dewc
                                          ON  dewc.EMERET_WASTE_ID = dwo.EMERET_WASTE_ID
                                    LEFT JOIN dbo.DATA_WASTE_CATEGORY dwc
                                          ON  dwc.WASTE_CODE = dwo.WASTE_CODE
                                    LEFT JOIN dbo.DATA_TREATMENT_DISTANCE td
                                          ON  td.CODE = dwo.LOCATION_CODE`;

  async getDataWMS() {
    const [dataLYV, dataLHG] = await Promise.all([
      this.LYV_WMS.query(this.querySQL, { type: QueryTypes.SELECT }),
      this.JAZ_WMS.query(this.querySQL, { type: QueryTypes.SELECT }),
    ]);

    const totalData = [...(dataLYV as any[]), ...(dataLHG as any[])]

    console.log(totalData);

    return 'The WMS data from Cat5 Service';
  }
}
