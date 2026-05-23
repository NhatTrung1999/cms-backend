import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { EventsGateway } from '../events/events.gateway';
import { v4 as uuidv4 } from 'uuid';
import * as path from 'path';
import * as fs from 'fs';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';
import { ConfigService } from '@nestjs/config';
import * as ExcelJS from 'exceljs';
import {
  FACTORY_LIST,
  FactoryCode,
  getFactory,
} from 'src/helper/factory.helper';
import { getCat6AirportName } from 'src/helper/cat6-airport.helper';
import {
  appendCat6DocumentSuffix,
  getCat6AssistedIds,
} from 'src/helper/cat6-assisted.helper';
import dayjs from 'dayjs';
import { ACTIVITY_TYPES, ActivityType } from 'src/types/cat1andcat4';
import {
  ACTIVITY_TYPES as ACTIVITY_TYPES_CAT9,
  ActivityType as ActivityTypeCat9,
} from 'src/types/cat9andcat12';
import {
  ACTIVITY_TYPES as ACTIVITY_TYPES_CAT5,
  ActivityType as ActivityTypeCat5,
} from 'src/types/cat5';
import {
  ACTIVITY_TYPES as ACTIVITY_TYPES_CAT7,
  ActivityType as ActivityTypeCat7,
} from 'src/types/cat7';
import { buildQueryAutoSentCMS } from 'src/helper/cat1andcat4.helper';
import { buildQueryAutoSentCMS as buildQueryAutoSentCMSCat9 } from 'src/helper/cat9andcat12.helper';
import { buildQueryAutoSentCMS as buildQueryAutoSentCMSCat5 } from 'src/helper/cat5.helper';
import {
  buildQueryAutoSentCmsLYV,
  buildQueryAutoSentCmsLHG,
  buildQueryAutoSentCmsLVL,
  buildQueryAutoSentCmsLYM,
  buildQueryAutoSentCmsJAZ,
  buildQueryAutoSentCmsJZS,
  buildReplacements,
} from 'src/helper/cat7.helper';
import { BuildQueryFn } from 'src/types/cat7';

type Cat6RouteItem = {
  AddressName: string;
  Transport: string;
  AddressDetail: string;
  isAirport: boolean;
  From: string;
  To: string;
};

type Cat6AccommodationItem = {
  nights?: number | string;
  addressID?: string;
  id?: string;
  isSameAsAbove?: boolean;
  type?: string;
};

type Cat6ActivePreviewRow = {
  Document_Date: string;
  Document_Number: string;
  Staff_ID: string;
  Dept: string;
  Round_trip_One_way: string;
  Start_Time: string;
  End_Time: string;
  Business_Trip_Type: string;
  Number_of_nights_stayed: number;
  Number_of_People: number;
  Factory_Code: string;
  [key: string]: string | number | undefined;
};

type Cat6PreviewPayload = {
  System: string;
  Corporation: string;
  Factory: string;
  Department: string;
  DocKey: string;
  ActivitySource: string;
  SPeriodData: string;
  EPeriodData: string;
  ActivityType: string;
  DataType: string;
  DocType: string;
  DocDate: string;
  DocDate2: string;
  DocNo: string;
  UndDocNo: string;
  TransType: string;
  Departure: string;
  Destination: string;
  Memo: string;
  CreateDateTime: string;
  Creator: string;
};

@Injectable()
export class PreviewpayloadService {
  private rootFolder: string;
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    private readonly eventsGateway: EventsGateway,
    private configService: ConfigService,
    //cat1&4, cat9&12
    @Inject('LYV_ERP') private readonly LYV_ERP: Sequelize,
    @Inject('LHG_ERP') private readonly LHG_ERP: Sequelize,
    @Inject('LVL_ERP') private readonly LVL_ERP: Sequelize,
    @Inject('LYM_ERP') private readonly LYM_ERP: Sequelize,
    @Inject('LYF_ERP') private readonly LYF_ERP: Sequelize,
    @Inject('JAZ_ERP') private readonly JAZ_ERP: Sequelize,
    @Inject('JZS_ERP') private readonly JZS_ERP: Sequelize,
    //cat5
    @Inject('LYV_WMS') private readonly LYV_WMS: Sequelize,
    @Inject('LHG_WMS') private readonly LHG_WMS: Sequelize,
    @Inject('LYM_WMS') private readonly LYM_WMS: Sequelize,
    @Inject('LVL_WMS') private readonly LVL_WMS: Sequelize,
    @Inject('JAZ_WMS') private readonly JAZ_WMS: Sequelize,
    @Inject('JZS_WMS') private readonly JZS_WMS: Sequelize,
    //cat7
    @Inject('LYV_HRIS') private readonly LYV_HRIS: Sequelize,
    @Inject('LHG_HRIS') private readonly LHG_HRIS: Sequelize,
    @Inject('LVL_HRIS') private readonly LVL_HRIS: Sequelize,
    @Inject('LYM_HRIS') private readonly LYM_HRIS: Sequelize,
    @Inject('JAZ_HRIS') private readonly JAZ_HRIS: Sequelize,
    @Inject('JZS_HRIS') private readonly JZS_HRIS: Sequelize,
    @Inject('UOF') private readonly UOF: Sequelize,
  ) {
    this.rootFolder = this.configService.get(
      'EXCEL_STORAGE_PATH',
      'D:/FileExcel',
    );
    if (!fs.existsSync(this.rootFolder)) {
      fs.mkdirSync(this.rootFolder, { recursive: true });
    }
  }

  private async insertFileRecord(
    id: string,
    module: string,
    fileName: string,
    filePath: string,
    factory: string,
    userID: string,
  ): Promise<void> {
    await this.EIP.query(
      `INSERT INTO CMW_File_Management
        (ID, Module, [File_Name], [Path], [Status], CreatedAt, CreatedFactory, CreatedDate)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      {
        replacements: [
          id,
          module,
          fileName,
          filePath,
          '0',
          userID,
          factory,
          new Date().toISOString(),
        ],
        type: QueryTypes.INSERT,
      },
    );
  }

  private async updateFileSuccess(id: string) {
    try {
      await this.EIP.query(
        `UPDATE CMW_File_Management SET [Status] = '1' WHERE ID = ?`,
        { replacements: [id], type: QueryTypes.UPDATE },
      );
      const result = await this.EIP.query(
        `SELECT * FROM CMW_File_Management WHERE ID = ? ORDER BY CreatedDate`,
        { replacements: [id], type: QueryTypes.SELECT },
      );
      return result;
    } catch (error) {
      throw new Error(error);
    }
  }

  private async updateFileError(id: string): Promise<void> {
    await this.EIP.query(
      `UPDATE CMW_File_Management SET [Status] = '2' WHERE ID = ?`,
      { replacements: [id], type: QueryTypes.UPDATE },
    );
  }

  private readonly buildQueryMap: Record<FactoryCode, BuildQueryFn> = {
    LYV: buildQueryAutoSentCmsLYV,
    LHG: buildQueryAutoSentCmsLHG,
    LVL: buildQueryAutoSentCmsLVL,
    LYM: buildQueryAutoSentCmsLYM,
    JAZ: buildQueryAutoSentCmsJAZ,
    JZS: buildQueryAutoSentCmsJZS,
  };

  //preview cat1 and cat4
  private getDbByFactory(factory: FactoryCode): Sequelize | null {
    const dbMap: Record<FactoryCode, Sequelize> = {
      LYV: this.LYV_ERP,
      LHG: this.LHG_ERP,
      LVL: this.LVL_ERP,
      LYM: this.LYM_ERP,
      LYF: this.LYF_ERP,
      JAZ: this.JAZ_ERP,
      JZS: this.JZS_ERP,
    };
    return dbMap[factory] ?? null;
  }

  private mapToCMSFormat(
    item: any,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const factory = item.FactoryCode ?? '';
    // const docKey = `${item.MatID ?? ''}${item.ReceivedNo ?? ''}`;
    const matId = item.MatID ?? '';
    const docKey = matId.substring(0, 1);
    if (docKey.startsWith('W')) {
      return [];
    }
    const docDate = item.RKDate ? dayjs(item.RKDate).format('YYYY/MM/DD') : '';
    // const docDate2 = item.PurDate
    //   ? dayjs(item.PurDate).format('YYYY/MM/DD')
    //   : '';
    // const docNo = item.PurNo ?? '';
    const custVenName = item.SupplierCode ?? '';
    const transportationMethod = item.TransportationMethod ?? '';
    const departure = item.Departure ?? '';
    const portOfDeparture = item.PortOfDeparture ?? '';
    const portOfArrival = item.PortOfArrival ?? '';
    const destination = item.Destination ?? '';
    const activityData = item.WeightUnitkg ?? 0;
    const qtyReceive = item.QtyReceive ?? 0;
    const receivedNo = item.ReceivedNo ?? '';

    let transType = '';
    let portType = '';

    switch (transportationMethod.trim().toLowerCase()) {
      case 'SEA'.trim().toLowerCase():
      case 'SEA+LAND'.trim().toLowerCase():
        transType = '海運';
        portType = '海港';
        break;
      case 'AIR'.trim().toLowerCase():
        transType = '空運';
        portType = '空港';
        break;
      case 'LAND'.trim().toLowerCase():
        transType = '陸運';
        portType = '';
        break;
      default:
        break;
    }

    return ACTIVITY_TYPES.filter((item) => item === dockeyCMS).map(
      (activityType: ActivityType) => ({
        System: 'CMS Web',
        Corporation: 'Lai Yih',
        Factory: factory,
        Department: '',
        DocKey:
          activityType.trim() === '4.1'
            ? `${activityType.trim()}${docKey}`
            : activityType.trim(),
        SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
        EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
        ActivityType: activityType.trim(),
        DataType: activityType.trim() === '3.1' ? '1' : '999',
        DocType: '入庫單',
        UndDoc: '',
        DocFlow: '',
        DocDate: docDate,
        DocDate2: docDate,
        DocNo: receivedNo.trim(),
        UndDocNo: '',
        CustVenName: custVenName,
        InvoiceNo: '',
        TransType: transType,
        Departure: departure,
        Destination: destination,
        PortType: portType,
        StPort: portOfDeparture,
        ThPort: '',
        EndPort: portOfArrival,
        Product: matId.trim(),
        Quity: qtyReceive,
        Amount: '',
        ActivityData: activityData,
        ActivityUnit: 'KG',
        Unit: '',
        UnitWeight: '',
        Memo: '',
        CreateDateTime: '',
        Creator: '',
        ActivitySource: matId.charAt(0),
      }),
    );
  }

  private async getCMSByFactory(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const db = this.getDbByFactory(factory);
    if (!db) return [];

    const query = await buildQueryAutoSentCMS(factory, this.EIP);

    const replacements =
      dateFrom && dateTo
        ? {
            startDate: dateFrom,
            endDate: dayjs(dateTo)
              .add(1, 'day')
              .startOf('day')
              .format('YYYY-MM-DD'),
          }
        : {};

    let data = await db.query<any>(query, {
      type: QueryTypes.SELECT,
      replacements,
    });

    data = await Promise.all(
      data.map(async (item: any) => {
        const taxFreeZone: any[] = await this.EIP.query(
          `SELECT TaxFreeZoneAddress
              FROM CMW_TAX_FREE_ZONE_ADDRESS
              WHERE SupplierID = ? AND Factory = ?`,
          {
            replacements: [item.SupplierCode.trim(), factory],
            type: QueryTypes.SELECT,
          },
        );
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

    return data.flatMap((item) =>
      this.mapToCMSFormat(item, dateFrom, dateTo, dockeyCMS),
    );
  }

  // cat9 & 12
  private mapToCMSFormatCat9(
    item: any,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const date = item.Date;
    const factory = item.Factory;
    const factoryAddress = item.Factory_address;
    const shipmentDate = item.Shipment_Date;
    const bookingNo = item.Booking_No;
    const customerID = item.Customer_ID;
    const invoiceNumber = item.Invoice_Number;
    const transportMethod = item.Transport_Method; // AIR OR SEA
    const portOfArrival = item.Port_Of_Arrival; // DEPARTURE AND END PORT
    const portOfDeparture = item.Port_Of_Departure; // ST PORT
    const quantity = item.Quantity;
    const grossWeight = item.Gross_Weight; // ACTIVITY DATA
    const createdDate = dayjs().format('YYYY-MM-DD HH:mm:ss');

    let dockey = '';
    let transtType = '';
    let portType = '';

    switch (transportMethod.trim().toLowerCase()) {
      case 'SEA'.trim().toLowerCase():
        dockey = '3.2.01';
        transtType = '海運';
        portType = '海港';
        break;
      case 'AIR'.trim().toLowerCase():
        dockey = '3.2.02';
        transtType = '空運';
        portType = '空港';
        break;
      case 'LAND'.trim().toLowerCase():
        dockey = '3.2.03';
        transtType = '陸運';
        portType = '';
        break;
      default:
        dockey = '';
        break;
    }

    return ACTIVITY_TYPES_CAT9.filter((item) => item === dockeyCMS).map(
      (activityType: ActivityTypeCat9) => ({
        System: 'ERP', // DEFAULT
        Corporation: 'LAI YIH', // DEFAULT
        Factory: factory,
        Department: 'Shipping', // DEFAULT
        // DocKey:
        //   transportMethod.trim().toLowerCase() !== 'AIR'.trim().toLowerCase()
        //     ? '3.2.01'
        //     : '3.2.02', // DEFAULT
        DocKey: activityType.trim() === '5.3'.trim() ? '5.3' : dockey, // DEFAULT
        SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
        EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
        ActivityType: activityType, // DEFAULT
        DataType: activityType.trim() === '3.2' ? '1' : '999', // DEFAULT
        DocType: 'Invoice', // DEFAULT 應收憑單
        UndDoc: '銷貨單', // DEFAULT
        DocFlow: '銷貨相關流程', // DEFAULT
        DocDate: dayjs(date).format('YYYY/MM/DD'),
        DocDate2: dayjs(shipmentDate).format('YYYY/MM/DD'),
        DocNo: bookingNo, // DEFAULT
        UndDocNo: '', // DEFAULT
        CustVenName: customerID,
        InvoiceNo: invoiceNumber,
        TransType: transtType,
        Departure: factoryAddress,
        Destination: portOfArrival,
        PortType: portType,
        StPort: portOfDeparture,
        ThPort: '', // DEFAULT
        EndPort: portOfArrival,
        Product: '', // DEFAULT
        Quity: quantity,
        Amount: '', // DEFAULT
        ActivityData: +Number(grossWeight).toFixed(2) || 0,
        ActivityUnit: 'Kg', // DEFAULT 噸
        Unit: '', // DEFAULT
        UnitWeight: '', // DEFAULT
        Memo: '', // DEFAULT
        CreateDateTime: createdDate,
        Creator: '',
        ActivitySource: 'Incineration (not energy recovery)',
      }),
    );
  }

  private async getCMSByFactoryCat9(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const db = this.getDbByFactory(factory);
    if (!db) return [];

    const query = await buildQueryAutoSentCMSCat9(
      dateFrom,
      dateTo,
      factory,
      this.EIP,
    );

    const replacements =
      dateFrom && dateTo ? [dateFrom, dateTo, dateFrom, dateTo] : [];

    let data = await db.query<any>(query, {
      type: QueryTypes.SELECT,
      replacements,
    });

    data = await Promise.all(
      data.map(async (item: any) => {
        let portOfDeparture = '';
        if (
          item?.Article_ID.trim().toLowerCase() ===
          'SAMPLE SHOE'.trim().toLowerCase()
        ) {
          item.Transport_Method = 'AIR';
          portOfDeparture = 'SGN';
        } else {
          switch (item?.Transport_Method.trim().toLowerCase()) {
            case '11 SC'.trim().toLowerCase():
            case 'Ocean'.trim().toLowerCase():
            case 'Sea Air Collect'.trim().toLowerCase():
            case 'SC'.trim().toLowerCase():
            case 'Sea/Air'.trim().toLowerCase():
            case 'Partial'.trim().toLowerCase():
              item.Transport_Method = 'SEA';
              break;
            case '20 CC'.trim().toLowerCase():
            case 'AC'.trim().toLowerCase():
            case '10 AC'.trim().toLowerCase():
            case 'CC'.trim().toLowerCase():
            case 'Air PP'.trim().toLowerCase():
            case 'Air Express'.trim().toLowerCase():
            case 'AE'.trim().toLowerCase():
              item.Transport_Method = 'AIR';
              break;
            case 'RAIL'.trim().toLowerCase():
            case 'Rail Collect'.trim().toLowerCase():
            case 'Rail Colle'.trim().toLowerCase():
              item.Transport_Method = 'Land';
              break;
            default:
              item.Transport_Method = 'SEA';
              break;
          }

          switch (item.Factory_Name.trim().toLowerCase()) {
            case 'LYV'.trim().toLowerCase():
            case 'LVL'.trim().toLowerCase():
            case 'LHG'.trim().toLowerCase():
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'SEA'.trim().toLowerCase()
              ) {
                portOfDeparture = 'VNCLP';
              }
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'Land'.trim().toLowerCase()
              ) {
                portOfDeparture = 'SOTH';
              }
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'AIR'.trim().toLowerCase()
              ) {
                portOfDeparture = 'SGN';
              }
              break;
            case 'LYM'.trim().toLowerCase():
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'SEA'.trim().toLowerCase()
              ) {
                portOfDeparture = 'MMRGN';
              } else {
                portOfDeparture = 'RGN';
              }
              break;
            case 'LYF'.trim().toLowerCase():
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'SEA'.trim().toLowerCase()
              ) {
                portOfDeparture = 'IDSRG';
              } else {
                portOfDeparture = 'CGK';
              }
              break;
            default:
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'SEA'.trim().toLowerCase()
              ) {
                portOfDeparture = 'VNCLP';
              }
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'Land'.trim().toLowerCase()
              ) {
                portOfDeparture = 'SOTH';
              }
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'AIR'.trim().toLowerCase()
              ) {
                portOfDeparture = 'SGN';
              }
              break;
          }
        }

        const queryPoA = `SELECT PortCode
                          FROM CMW_PortCode
                          WHERE TransportMethod LIKE ? AND CustomerNumber LIKE ?`;
        const result: any = await this.EIP.query(queryPoA, {
          type: QueryTypes.SELECT,
          replacements: [`%${item.Transport_Method}%`, `%${item.Customer_ID}%`],
        });
        // console.log(result[0].PortCode);
        return {
          ...item,
          Transport_Method: item.Transport_Method,
          Port_Of_Departure: portOfDeparture,
          Port_Of_Arrival: result[0]?.PortCode.trim().toUpperCase() || 'N/A',
        };
      }),
    );
    return data.flatMap((item) =>
      this.mapToCMSFormatCat9(item, dateFrom, dateTo, dockeyCMS),
    );
  }

  // cat5
  private getDbByFactoryCat5(factory: FactoryCode): Sequelize | null {
    const dbMap: Record<string, Sequelize> = {
      LYV: this.LYV_WMS,
      LHG: this.LHG_WMS,
      LVL: this.LVL_WMS,
      LYM: this.LYM_WMS,
      JAZ: this.JAZ_WMS,
      JZS: this.JZS_WMS,
    };
    return dbMap[factory] ?? null;
  }
  private mapToCMSFormatCat5(
    item: any,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const custVenName = item.Vendor_ID;
    const departure = item.Factory_address;
    const destination = item.Waste_collection_address;
    const product = `${item.The_type_of_waste} - ${item.Waste_type}`;
    const activityData = item.Weight_of_waste_treated_Unit_kg;
    const wasteTreatmentMethod = item.Waste_Treatment_method;
    const factoryName = item.Factory_Name;
    const wasteDisposalDate = item.Waste_disposal_date;
    // const wasteCode = item.Waste_Code;
    let dockey = '';

    switch (wasteTreatmentMethod.trim().toLowerCase()) {
      case 'Waste Recycling'.trim().toLowerCase():
        dockey = '4.4.01';
        break;
      case 'Waste to Energy'.trim().toLowerCase():
        dockey = '4.4.02';
        break;
      case 'Landfill'.trim().toLowerCase():
        dockey = '4.4.03';
        break;
      case 'Waste Incineration'.trim().toLowerCase():
        dockey = '4.4.04';
        break;
      case 'Waste other treatment method'.trim().toLowerCase():
        dockey = '4.4.05';
        break;
      default:
        dockey = '';
        break;
    }

    return ACTIVITY_TYPES_CAT5.filter((item) => item === dockeyCMS).map(
      (activityType: ActivityTypeCat5) => ({
        System: 'CMW', //default
        Corporation: 'LAI YIH', //default
        Factory: factoryName,
        Department: '',
        // DocKey: `${dayjs(wasteDisposalDate).format('YYYY/MM/DD')} ${wasteCode}`, //default
        DocKey: activityType.trim() === '3.6' ? '3.6' : dockey, //default
        SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
        EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
        ActivityType: activityType, //default
        DataType: activityType.trim() === '3.6' ? '1' : '999', //default
        DocType: 'CMS Web', //default
        UndDoc: '',
        DocFlow: '',
        DocDate: dayjs(wasteDisposalDate).format('YYYY/MM/DD'),
        DocDate2: dayjs(wasteDisposalDate).format('YYYY/MM/DD'),
        DocNo: '',
        UndDocNo: '',
        CustVenName: custVenName,
        InvoiceNo: '',
        TransType: '陸運', //default
        Departure: departure,
        Destination: destination,
        PortType: '',
        StPort: '',
        ThPort: '',
        EndPort: '',
        Product: product,
        Quity: '',
        Amount: '',
        ActivityData: activityData,
        ActivityUnit: 'KG', //default
        Unit: '',
        UnitWeight: '',
        Memo: '',
        CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
        Creator: '',
        AccNo: '6002',
        AccName: '其他費用',
        ActivitySource: wasteTreatmentMethod,
      }),
    );
  }
  private async getCMSByFactoryCat5(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const db = this.getDbByFactoryCat5(factory);
    if (!db) return [];

    const query = await buildQueryAutoSentCMSCat5(
      dateFrom,
      dateTo,
      factory,
      this.EIP,
    );

    const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

    const data = await db.query<any>(query, {
      type: QueryTypes.SELECT,
      replacements,
    });

    return data.flatMap((item) =>
      this.mapToCMSFormatCat5(item, dateFrom, dateTo, dockeyCMS),
    );
  }

  // cat7
  private getDbByFactoryCat7(factory: FactoryCode): Sequelize | null {
    const dbMap: Record<FactoryCode, Sequelize> = {
      LYV: this.LYV_HRIS,
      LHG: this.LHG_HRIS,
      LVL: this.LVL_HRIS,
      LYM: this.LYM_HRIS,
      JAZ: this.JAZ_HRIS,
      JZS: this.JZS_HRIS,
    };

    return dbMap[factory] ?? null;
  }
  private mapToCMSFormatCat7(
    item: any,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const staffId = item.Staff_ID ?? '';
    const Main_transportation_type = item.Main_transportation_type ?? '';
    const Number_of_working_days = item.Number_of_working_days ?? '';
    const Factory_address = item.Factory_address ?? '';
    const Residential_address = item.Residential_address ?? Factory_address;
    const FactoryName = item.Factory_Name ?? '';
    const DepartmentName = item.Department_Name ?? '';

    // if (DepartmentName.length > 20) {
    //   return [];
    // }

    // console.log(DepartmentName.length);
    let dockey = '';
    let transType = '';

    switch (Main_transportation_type.trim().toLowerCase()) {
      case 'Walking'.trim().toLowerCase(): //步行
        dockey = '3.3.1';
        transType = '步行';
        break;
      case 'Bicycle'.trim().toLowerCase(): //自行車
        dockey = '3.3.2';
        transType = '自行車';
        break;
      case 'Electric motorcycle'.trim().toLowerCase(): //電動機車
        dockey = '3.3.3';
        transType = '電動機車';
        break;
      case 'Motorcycle'.trim().toLowerCase(): //機車
        dockey = '3.3.4';
        transType = '機車';
        break;
      case 'Electric car'.trim().toLowerCase(): //電動汽車
        dockey = '3.3.5';
        transType = '電動汽車';
        break;
      case 'Car'.trim().toLowerCase(): //汽車
        dockey = '3.3.6';
        transType = '汽車';
        break;
      case 'Bus'.trim().toLowerCase(): //公車/客運
        dockey = '3.3.7';
        transType = '公車/客運';
        break;
      case 'Company shuttle bus'.trim().toLowerCase(): //員工接駁車
        dockey = '3.3.8';
        transType = '員工接駁車';
        break;
      case 'Subway'.trim().toLowerCase(): //地鐵
        dockey = '3.3.9';
        transType = '地鐵';
        break;
      default:
        break;
    }

    return ACTIVITY_TYPES_CAT7.filter((item) => item === dockeyCMS).map(
      (activityType: ActivityTypeCat7) => ({
        // System: 'CMS Web', // Default
        // Corporation: 'LAI YIH', // Default
        // Factory: FactoryName,
        // Department: DepartmentName,
        // DocKey: staffId,
        // SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
        // EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
        // ActivityType: activityType, // Default
        // DataType: '2', // Default
        // DocType: '員工通勤', // Default
        // DocDate: dayjs().format('YYYY/MM/DD'),
        // DocDate2: dayjs().format('YYYY/MM/DD'),
        // UndDocNo: staffId,
        // TransType: Main_transportation_type,
        // Departure: Factory_address ? Factory_address : '',
        // Destination: Residential_address ? Residential_address : '',
        // Attendance: Number_of_working_days.toString(),
        // Memo: '',
        // CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
        // Creator: '',
        System: 'CMS Web',
        Corporation: 'LAI YIH',
        Factory: FactoryName,
        Department: DepartmentName.slice(0, 20),
        DocKey: dockey,
        SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
        EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
        ActivityType: activityType,
        DataType: '2',
        DocType: 'CMS Web',
        DocDate: dayjs('2025/12/31').format('YYYY/MM/DD'),
        DocDate2: dayjs('2025/12/31').format('YYYY/MM/DD'),
        UndDocNo: staffId,
        TransType: transType,
        Departure: Residential_address,
        Destination: Factory_address,
        Attendance: Number_of_working_days.toString(),
        Memo: '',
        CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
        Creator: '',
      }),
    );
  }

  private async getCMSByFactoryCat7(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const db = this.getDbByFactoryCat7(factory);
    if (!db) return [];

    const buildQuery = this.buildQueryMap[factory];
    if (!buildQuery) return [];

    const { dateToExclusive, hasDate } = buildReplacements(dateFrom, dateTo);

    const query = await buildQuery(dateFrom, dateTo, this.EIP);

    const replacements = hasDate ? [dateFrom, dateToExclusive] : [];

    const data = await db.query<any>(query, {
      type: QueryTypes.SELECT,
      replacements,
    });

    return data.flatMap((item) =>
      this.mapToCMSFormatCat7(item, dateFrom, dateTo, dockeyCMS),
    );
  }

  private parseJsonArray<T = unknown>(value: unknown): T[] {
    if (Array.isArray(value)) return value as T[];
    if (typeof value !== 'string' || !value.trim()) return [];
    try {
      const parsed = JSON.parse(value);
      return Array.isArray(parsed) ? (parsed as T[]) : [];
    } catch {
      return [];
    }
  }

  private formatCat6TripType(type: unknown): string | null {
    if (!type || typeof type !== 'string') return null;
    const t = type.toLowerCase().trim();
    if (t === 'round' || t === 'roundtrip' || t === 'round trip') {
      return 'Round trip';
    }
    return 'One-way';
  }

  private formatCat6BusinessTripType(type: unknown): string | null {
    if (!type || typeof type !== 'string') return null;
    const map: Record<string, string> = {
      factory_domestic: 'Domestic business trip within the group',
      outside_domestic: 'Domestic business trip to third party entities',
      factory_oversea: 'Overseas business trip within the group',
      outside_oversea: 'Overseas business trip to third party entities',
    };
    return map[type.toLowerCase()] ?? type;
  }

  private formatCat6PlacesAndTransports(routesValue: unknown) {
    const routes = this.parseJsonArray<Cat6RouteItem>(routesValue);
    const places: string[] = [];
    const transports: string[] = [];

    const pushPlace = (value?: string) => {
      const place = getCat6AirportName(value);
      if (!place) return;
      if (places[places.length - 1] === place) return;
      places.push(place);
    };

    const isFlightRoute = (route: Cat6RouteItem) => {
      const transport = route.Transport?.trim().toLowerCase() ?? '';
      return route.isAirport === true || transport === 'flight';
    };

    let index = 0;
    while (index < routes.length) {
      const route = routes[index];
      const transport = route.Transport?.trim() ?? '';
      const nextRoute = routes[index + 1];

      if (isFlightRoute(route)) {
        const flightStart = route.From?.trim() ?? '';
        let flightEnd = route.To?.trim() ?? '';
        let cursor = index;

        while (cursor + 1 < routes.length && isFlightRoute(routes[cursor + 1])) {
          cursor += 1;
          flightEnd = routes[cursor].To?.trim() ?? flightEnd;
        }

        pushPlace(flightStart);
        pushPlace(flightEnd);
        transports.push('Flight');
        index = cursor + 1;
        continue;
      }

      if (!transport && places.length === 0 && nextRoute && isFlightRoute(nextRoute)) {
        pushPlace(nextRoute.From);
        index += 1;
        continue;
      }

      pushPlace(route.AddressDetail ?? route.AddressName);
      if (transport) {
        transports.push(transport);
      }
      index += 1;
    }

    const transportLimit = Math.max(places.length - 1, 0);
    const normalizedTransports = transports.slice(0, transportLimit);
    while (normalizedTransports.length < transportLimit) {
      normalizedTransports.push('');
    }

    return {
      ...places.reduce<Record<string, string>>((acc, place, index) => {
        acc[`Place${index + 1}`] = place;
        return acc;
      }, {}),
      ...normalizedTransports.reduce<Record<string, string>>(
        (acc, transport, index) => {
          acc[`Transport_${index + 1}`] = transport;
          return acc;
        },
        {},
      ),
    };
  }

  private getCat6AccommodationNights(accommodationValue: unknown) {
    const accommodations =
      this.parseJsonArray<Cat6AccommodationItem>(accommodationValue);
    return accommodations.reduce((sum, item) => {
      const nights = Number(item?.nights ?? 0);
      return sum + (Number.isFinite(nights) ? nights : 0);
    }, 0);
  }

  private getCat6NumberOfPeople(assistedIdsValue: unknown) {
    return getCat6AssistedIds(assistedIdsValue).length;
  }

  private expandCat6PreviewRowsByAssistedIds(
    row: Record<string, any>,
  ): Record<string, any>[] {
    const assistedIds = getCat6AssistedIds(row.AssisstedIDs);

    if (assistedIds.length === 0) {
      return [
        {
          ...row,
          DOC_NBR: row.DOC_NBR ?? '',
          EffectiveStaffID: row.UserCreate ?? '',
          Number_of_People: 0,
        },
      ];
    }

    if (assistedIds.length === 1) {
      return [
        {
          ...row,
          DOC_NBR: row.DOC_NBR ?? '',
          EffectiveStaffID: assistedIds[0],
          Number_of_People: 1,
        },
      ];
    }

    return assistedIds.map((assistedId, index) => ({
      ...row,
      DOC_NBR: appendCat6DocumentSuffix(String(row.DOC_NBR ?? ''), index),
      EffectiveStaffID: assistedId,
      Number_of_People: 1,
    }));
  }

  private transformCat6ActiveRow(
    row: Record<string, any>,
  ): Cat6ActivePreviewRow {
    return {
      Document_Date: row.CreatedAt
        ? dayjs(row.CreatedAt).format('YYYY-MM-DD')
        : '',
      Document_Number: row.DOC_NBR ?? '',
      Staff_ID: row.EffectiveStaffID ?? '',
      Dept: row.Dept ?? row.Department ?? '',
      Round_trip_One_way: this.formatCat6TripType(row.TypeTravel) ?? '',
      Start_Time: row.DateStart ? dayjs(row.DateStart).format('YYYY-MM-DD') : '',
      End_Time: row.DateEnd ? dayjs(row.DateEnd).format('YYYY-MM-DD') : '',
      Business_Trip_Type: this.formatCat6BusinessTripType(row.Factory) ?? '',
      ...this.formatCat6PlacesAndTransports(row.Routes),
      Number_of_nights_stayed: this.getCat6AccommodationNights(row.Accommodation),
      Number_of_People: Number(row.Number_of_People ?? 0),
      Factory_Code: String(row.Factory_User ?? '').trim().toUpperCase(),
    };
  }

  private extractCat6Places(row: Cat6ActivePreviewRow): string[] {
    return Object.keys(row)
      .filter((key) => /^Place\d+$/.test(key))
      .sort((left, right) => {
        const leftIndex = Number(left.replace('Place', ''));
        const rightIndex = Number(right.replace('Place', ''));
        return leftIndex - rightIndex;
      })
      .map((key) => String(row[key] ?? '').trim())
      .filter(Boolean);
  }

  private extractCat6Transports(row: Cat6ActivePreviewRow): string[] {
    return Object.keys(row)
      .filter((key) => /^Transport_\d+$/.test(key))
      .sort((left, right) => {
        const leftIndex = Number(left.replace('Transport_', ''));
        const rightIndex = Number(right.replace('Transport_', ''));
        return leftIndex - rightIndex;
      })
      .map((key) => String(row[key] ?? '').trim());
  }

  private buildCat6CMSLegs(row: Cat6ActivePreviewRow) {
    const places = this.extractCat6Places(row);
    const transports = this.extractCat6Transports(row);
    const legs: { transType: string; dep: string; dest: string }[] = [];
    const maxLegs = Math.max(places.length - 1, 0);

    for (let index = 0; index < maxLegs; index += 1) {
      const dep = places[index]?.trim() ?? '';
      const dest = places[index + 1]?.trim() ?? '';
      const transType = transports[index]?.trim() ?? '';

      if (!dep || !dest || !transType) {
        continue;
      }

      legs.push({
        transType,
        dep,
        dest,
      });
    }

    return legs;
  }

  private getCat6CMSDocKey(businessTripType: string): string {
    switch (businessTripType.trim().toLowerCase()) {
      case 'domestic business trip within the group':
        return '3.5.1';
      case 'domestic business trip to third party entities':
        return '3.5.2';
      case 'overseas business trip within the group':
        return '3.5.3';
      case 'overseas business trip to third party entities':
        return '3.5.4';
      default:
        return '';
    }
  }

  private mapCat6CMSTransType(transport: string): string {
    const normalized = transport.trim().toLowerCase();

    switch (normalized) {
      case 'car':
        return '計程車/出租車';
      case 'company shuttle car':
        return '公司車';
      case 'flight':
        return '飛機';
      case 'personal motocycle':
        return '個人機車';
      default:
        return transport.trim();
    }
  }

  private mapCat6RowToPreviewPayload(
    row: Cat6ActivePreviewRow,
    factory: string,
    dateFrom: string,
    dateTo: string,
  ): Cat6PreviewPayload[] {
    const legs = this.buildCat6CMSLegs(row);
    const docKey = this.getCat6CMSDocKey(row.Business_Trip_Type ?? '');
    const resolvedFactoryCode =
      factory.trim().toUpperCase() === 'ALL'
        ? row.Factory_Code
        : factory.trim().toUpperCase();

    return legs.map((leg, index) => ({
      System: 'BPM',
      Corporation: 'LAI YIH',
      Factory: getFactory(resolvedFactoryCode as FactoryCode),
      Department: row.Dept ?? '',
      DocKey: docKey,
      ActivitySource: '',
      SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
      EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
      ActivityType: '3.5',
      DataType: '2',
      DocType: '洽公單',
      DocDate: row.Document_Date
        ? dayjs(row.Document_Date).format('YYYY/MM/DD')
        : '',
      DocDate2: row.Start_Time
        ? dayjs(row.Start_Time).format('YYYY/MM/DD')
        : '',
      DocNo: row.Document_Number ?? '',
      UndDocNo: `${row.Document_Number ?? ''}-${index + 1}`,
      TransType: this.mapCat6CMSTransType(leg.transType),
      Departure: leg.dep,
      Destination: leg.dest,
      Memo: row.Business_Trip_Type ?? '',
      CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
      Creator: '',
    }));
  }

  async exportExcelCat6(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    let where = `WHERE  chb.BPMStatus = 'F'
                        AND ISNULL(chb.Factory_User ,'')<>''`;
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ` AND CONVERT(VARCHAR ,chb.CreatedAt ,23) BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }

    if (factory && factory.trim().toUpperCase() !== 'ALL') {
      where += ` AND chb.Factory_User LIKE ?`;
      replacements.push(`%${factory}%`);
    }

    const query = `
      SELECT chb.*
                          --,twt.Current_doc
                          ,COALESCE(
                              vwd.GROUP_NAME
                              ,(
                                  SELECT TOP 1 vwd2.GROUP_NAME
                                  FROM   TB_EB_USER teu2
                                          OUTER APPLY (
                                      SELECT teed2.GROUP_ID
                                      FROM   TB_EB_EMPL_DEP AS teed2
                                      WHERE  teed2.USER_GUID = teu2.USER_GUID
                                              AND teed2.ORDERS = 0
                                  ) teed2
                                  LEFT JOIN vwDepartment_Factory vwd2
                                              ON  vwd2.GROUP_ID = teed2.GROUP_ID
                                  WHERE  teu2.ACCOUNT = chb.Factory_User+chb.UserCreate
                                          AND vwd2.GROUP_NAME IS NOT NULL
                              )
                              ,(
                                  SELECT TOP 1 vwd3.GROUP_NAME
                                  FROM   TB_EB_USER teu3
                                          OUTER APPLY (
                                      SELECT teed3.GROUP_ID
                                      FROM   TB_EB_EMPL_DEP AS teed3
                                      WHERE  teed3.USER_GUID = teu3.USER_GUID
                                              AND teed3.ORDERS = 0
                                  ) teed3
                                  LEFT JOIN vwDepartment_Factory vwd3
                                              ON  vwd3.GROUP_ID = teed3.GROUP_ID
                                  WHERE  teu3.ACCOUNT = chb.Departure+chb.UserCreate
                                          AND vwd3.GROUP_NAME IS NOT NULL
                              )
                              ,(
                                  SELECT TOP 1 vwd4.GROUP_NAME
                                  FROM   CDS_FMEval_Employee cfe
                                          JOIN TB_EB_USER teu4
                                              ON  teu4.ACCOUNT = cfe.BPMAccount
                                          OUTER APPLY (
                                      SELECT teed4.GROUP_ID
                                      FROM   TB_EB_EMPL_DEP AS teed4
                                      WHERE  teed4.USER_GUID = teu4.USER_GUID
                                              AND teed4.ORDERS = 0
                                  ) teed4
                                  LEFT JOIN vwDepartment_Factory vwd4
                                              ON  vwd4.GROUP_ID = teed4.GROUP_ID
                                  WHERE  cfe.EmpID = chb.UserCreate
                                          AND vwd4.GROUP_NAME IS NOT NULL
                              )
                          )                         AS Dept
                          ,COUNT(chb.TripID) OVER()  AS TotalRow
                    FROM   CDS_HRBUSS_BusTripData    AS chb
                          LEFT JOIN TB_EB_USER      AS teu
                                ON  (
                                        teu.ACCOUNT LIKE chb.Factory_User+'[0-9][0-9][0-9][0-9][0-9]'
                                        AND teu.ACCOUNT LIKE '%'+chb.UserCreate+'%'
                                    )
                                    OR (
                                        teu.ACCOUNT NOT LIKE chb.Factory_User+'[0-9][0-9][0-9][0-9][0-9]'
                                        AND teu.ACCOUNT=chb.UserCreate
                                    )
                          OUTER APPLY (
                        SELECT teed2.GROUP_ID
                        FROM   TB_EB_EMPL_DEP AS teed2
                        WHERE  teed2.USER_GUID = teu.USER_GUID
                              AND teed2.ORDERS = 0
                    )                                AS teed
                    LEFT JOIN vwDepartment_Factory   AS vwd
                                ON  vwd.GROUP_ID = teed.GROUP_ID
                          LEFT JOIN tb_wkf_task     AS twt
                                ON  twt.DOC_NBR = chb.DOC_NBR
                                    AND (
                                            twt.DOC_NBR LIKE 'LYV-HR-BT%'
                                            OR twt.DOC_NBR LIKE 'LHG-SUGG%'
                                            OR twt.DOC_NBR LIKE 'LVL-HR-BTF%'
                                            OR twt.DOC_NBR LIKE 'LVL-ODBT%'
                                            OR twt.DOC_NBR LIKE 'LYM-HR-BT%'
                                        )
      ${where}
      ORDER BY CreatedAt ASC
    `;

    const rawRows = (await this.UOF.query(query, {
      type: QueryTypes.SELECT,
      replacements,
    })) as Record<string, any>[];

    const transformedRows = rawRows.flatMap((row) =>
      this.expandCat6PreviewRowsByAssistedIds(row).map((expandedRow) =>
        this.transformCat6ActiveRow(expandedRow),
      ),
    );
    const payloadRows = transformedRows.flatMap((row) =>
      this.mapCat6RowToPreviewPayload(row, factory, dateFrom, dateTo),
    );

    sheet.columns = [
      { header: 'System', key: 'System', width: 15 },
      { header: 'Corporation', key: 'Corporation', width: 15 },
      { header: 'Factory', key: 'Factory', width: 18 },
      { header: 'Department', key: 'Department', width: 20 },
      { header: 'DocKey', key: 'DocKey', width: 12 },
      { header: 'ActivitySource', key: 'ActivitySource', width: 15 },
      { header: 'SPeriodData', key: 'SPeriodData', width: 15 },
      { header: 'EPeriodData', key: 'EPeriodData', width: 15 },
      { header: 'ActivityType', key: 'ActivityType', width: 12 },
      { header: 'DataType', key: 'DataType', width: 10 },
      { header: 'DocType', key: 'DocType', width: 18 },
      { header: 'DocDate', key: 'DocDate', width: 15 },
      { header: 'DocDate2', key: 'DocDate2', width: 15 },
      { header: 'DocNo', key: 'DocNo', width: 20 },
      { header: 'UndDocNo', key: 'UndDocNo', width: 22 },
      { header: 'TransType', key: 'TransType', width: 20 },
      { header: 'Departure', key: 'Departure', width: 35 },
      { header: 'Destination', key: 'Destination', width: 35 },
      { header: 'Memo', key: 'Memo', width: 35 },
      { header: 'CreateDateTime', key: 'CreateDateTime', width: 20 },
      { header: 'Creator', key: 'Creator', width: 15 },
    ];

    sheet.getRow(1).font = { bold: true };

    payloadRows.forEach((item) => sheet.addRow(item));
  }
  // Preview Excel
  async previewExcel(
    module: string,
    dateFrom: string,
    dateTo: string,
    factory: string,
    userID: string,
    dockeyCMS: string,
  ) {
    const id = uuidv4();
    const moduleName = module.trim();
    const fileName = `Preview-Payload-${factory}-${moduleName}-${dateFrom}-${dateTo}-${dockeyCMS}.xlsx`;
    const folder = path.join(this.rootFolder, moduleName);
    if (!fs.existsSync(folder)) fs.mkdirSync(folder, { recursive: true });
    const filePath = path.join(folder, fileName);

    await this.insertFileRecord(
      id,
      moduleName,
      fileName,
      filePath,
      factory,
      userID,
    );

    await this.createFileExcel(
      id,
      moduleName,
      filePath,
      dateFrom,
      dateTo,
      factory,
      dockeyCMS,
    );

    return true;
  }

  async createFileExcel(
    id: string,
    module: string,
    filePath: string,
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(module);

    switch (module.trim().toLowerCase()) {
      case 'cat1andcat4':
        await this.exportExcelCat1And4(
          sheet,
          dateFrom,
          dateTo,
          factory,
          dockeyCMS,
        );
        break;
      case 'cat5':
        await this.exportExcelCat5(sheet, dateFrom, dateTo, factory, dockeyCMS);
        break;
      case 'cat6':
        await this.exportExcelCat6(sheet, dateFrom, dateTo, factory, dockeyCMS);
        break;
      case 'cat7':
        await this.exportExcelCat7(sheet, dateFrom, dateTo, factory, dockeyCMS);
        break;
      case 'cat9andcat12':
        await this.exportExcelCat9And12(
          sheet,
          dateFrom,
          dateTo,
          factory,
          dockeyCMS,
        );
        break;
      default:
        throw new InternalServerErrorException(
          `Module "${module}" không hợp lệ`,
        );
    }

    await workbook.xlsx
      .writeFile(filePath)
      .then(async () => {
        setTimeout(async () => {
          const statusDone = await this.updateFileSuccess(id);
          if (statusDone.length) {
            this.eventsGateway.broadcastEvent(
              'file-excel-done',
              'Export file excel success!',
            );
          }
        }, 60000);
      })
      .catch(async (error: any) => {
        await this.updateFileError(id);
        this.eventsGateway.broadcastEvent(
          'file-excel-error',
          'Error export file excel!',
        );
      });
  }

  async exportExcelCat1And4(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    let data: any[];
    if (factory.trim().toUpperCase() === 'ALL') {
      const results = await Promise.all(
        FACTORY_LIST.map((f) =>
          this.getCMSByFactory(f, dateFrom, dateTo, dockeyCMS),
        ),
      );
      data = results.flat();
    } else {
      data = await this.getCMSByFactory(
        factory as FactoryCode,
        dateFrom,
        dateTo,
        dockeyCMS,
      );
    }

    // Set columns giống exportPreviewPayload
    sheet.columns = [
      { header: 'System', key: 'System', width: 15 },
      { header: 'Corporation', key: 'Corporation', width: 15 },
      { header: 'Factory', key: 'Factory', width: 15 },
      { header: 'Department', key: 'Department', width: 15 },
      { header: 'DocKey', key: 'DocKey', width: 15 },
      { header: 'SPeriodData', key: 'SPeriodData', width: 15 },
      { header: 'EPeriodData', key: 'EPeriodData', width: 15 },
      { header: 'ActivityType', key: 'ActivityType', width: 15 },
      { header: 'DataType', key: 'DataType', width: 15 },
      { header: 'DocType', key: 'DocType', width: 15 },
      { header: 'UndDoc', key: 'UndDoc', width: 15 },
      { header: 'DocFlow', key: 'DocFlow', width: 15 },
      { header: 'DocDate', key: 'DocDate', width: 15 },
      { header: 'DocDate2', key: 'DocDate2', width: 15 },
      { header: 'DocNo', key: 'DocNo', width: 15 },
      { header: 'UndDocNo', key: 'UndDocNo', width: 15 },
      { header: 'CustVenName', key: 'CustVenName', width: 25 },
      { header: 'InvoiceNo', key: 'InvoiceNo', width: 15 },
      { header: 'TransType', key: 'TransType', width: 15 },
      { header: 'Departure', key: 'Departure', width: 25 },
      { header: 'Destination', key: 'Destination', width: 15 },
      { header: 'PortType', key: 'PortType', width: 15 },
      { header: 'StPort', key: 'StPort', width: 15 },
      { header: 'ThPort', key: 'ThPort', width: 15 },
      { header: 'EndPort', key: 'EndPort', width: 15 },
      { header: 'Product', key: 'Product', width: 20 },
      { header: 'Quity', key: 'Quity', width: 12 },
      { header: 'Amount', key: 'Amount', width: 12 },
      { header: 'ActivityData', key: 'ActivityData', width: 15 },
      { header: 'ActivityUnit', key: 'ActivityUnit', width: 15 },
      { header: 'Unit', key: 'Unit', width: 10 },
      { header: 'UnitWeight', key: 'UnitWeight', width: 12 },
      { header: 'Memo', key: 'Memo', width: 20 },
      { header: 'CreateDateTime', key: 'CreateDateTime', width: 20 },
      { header: 'Creator', key: 'Creator', width: 15 },
      { header: 'ActivitySource', key: 'ActivitySource', width: 20 },
    ];

    sheet.getRow(1).font = { bold: true };

    for (const row of data) {
      sheet.addRow(row);
    }
  }

  async exportExcelCat5(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    let data: any[];
    if (factory.trim().toUpperCase() === 'ALL') {
      const results = await Promise.all(
        FACTORY_LIST.map((f) =>
          this.getCMSByFactoryCat5(f, dateFrom, dateTo, dockeyCMS),
        ),
      );
      data = results.flat();
    } else {
      data = await this.getCMSByFactoryCat5(
        factory as FactoryCode,
        dateFrom,
        dateTo,
        dockeyCMS,
      );
    }

    sheet.columns = [
      { header: 'System', key: 'System', width: 15 },
      { header: 'Corporation', key: 'Corporation', width: 15 },
      { header: 'Factory', key: 'Factory', width: 15 },
      { header: 'Department', key: 'Department', width: 15 },
      { header: 'DocKey', key: 'DocKey', width: 15 },
      { header: 'SPeriodData', key: 'SPeriodData', width: 15 },
      { header: 'EPeriodData', key: 'EPeriodData', width: 15 },
      { header: 'ActivityType', key: 'ActivityType', width: 15 },
      { header: 'DataType', key: 'DataType', width: 15 },
      { header: 'DocType', key: 'DocType', width: 15 },
      { header: 'UndDoc', key: 'UndDoc', width: 15 },
      { header: 'DocFlow', key: 'DocFlow', width: 15 },
      { header: 'DocDate', key: 'DocDate', width: 15 },
      { header: 'DocDate2', key: 'DocDate2', width: 15 },
      { header: 'DocNo', key: 'DocNo', width: 15 },
      { header: 'UndDocNo', key: 'UndDocNo', width: 15 },
      { header: 'CustVenName', key: 'CustVenName', width: 25 },
      { header: 'InvoiceNo', key: 'InvoiceNo', width: 15 },
      { header: 'TransType', key: 'TransType', width: 15 },
      { header: 'Departure', key: 'Departure', width: 25 },
      { header: 'Destination', key: 'Destination', width: 25 },
      { header: 'PortType', key: 'PortType', width: 15 },
      { header: 'StPort', key: 'StPort', width: 15 },
      { header: 'ThPort', key: 'ThPort', width: 15 },
      { header: 'EndPort', key: 'EndPort', width: 15 },
      { header: 'Product', key: 'Product', width: 30 },
      { header: 'Quity', key: 'Quity', width: 12 },
      { header: 'Amount', key: 'Amount', width: 12 },
      { header: 'ActivityData', key: 'ActivityData', width: 15 },
      { header: 'ActivityUnit', key: 'ActivityUnit', width: 15 },
      { header: 'Unit', key: 'Unit', width: 10 },
      { header: 'UnitWeight', key: 'UnitWeight', width: 12 },
      { header: 'Memo', key: 'Memo', width: 20 },
      { header: 'CreateDateTime', key: 'CreateDateTime', width: 20 },
      { header: 'Creator', key: 'Creator', width: 15 },
      { header: 'AccNo', key: 'AccNo', width: 10 },
      { header: 'AccName', key: 'AccName', width: 15 },
      { header: 'ActivitySource', key: 'ActivitySource', width: 25 },
    ];

    sheet.getRow(1).font = { bold: true };

    for (const row of data) {
      sheet.addRow(row);
    }
  }

  async exportExcelCat7(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    let data: any[];
    if (factory.trim().toUpperCase() === 'ALL') {
      const results = await Promise.all(
        FACTORY_LIST.map((f) => this.getCMSByFactoryCat7(f, dateFrom, dateTo, dockeyCMS)),
      );
      data = results.flat();
    } else {
      data = await this.getCMSByFactoryCat7(factory, dateFrom, dateTo, dockeyCMS);
    }

    sheet.columns = [
      { header: 'System', key: 'System', width: 15 },
      { header: 'Corporation', key: 'Corporation', width: 15 },
      { header: 'Factory', key: 'Factory', width: 15 },
      { header: 'Department', key: 'Department', width: 20 },
      { header: 'DocKey', key: 'DocKey', width: 15 },
      { header: 'SPeriodData', key: 'SPeriodData', width: 15 },
      { header: 'EPeriodData', key: 'EPeriodData', width: 15 },
      { header: 'ActivityType', key: 'ActivityType', width: 15 },
      { header: 'DataType', key: 'DataType', width: 15 },
      { header: 'DocType', key: 'DocType', width: 15 },
      { header: 'DocDate', key: 'DocDate', width: 15 },
      { header: 'DocDate2', key: 'DocDate2', width: 15 },
      { header: 'UndDocNo', key: 'UndDocNo', width: 15 },
      { header: 'TransType', key: 'TransType', width: 15 },
      { header: 'Departure', key: 'Departure', width: 25 },
      { header: 'Destination', key: 'Destination', width: 25 },
      { header: 'Attendance', key: 'Attendance', width: 15 },
      { header: 'Memo', key: 'Memo', width: 20 },
      { header: 'CreateDateTime', key: 'CreateDateTime', width: 20 },
      { header: 'Creator', key: 'Creator', width: 15 },
    ];

    sheet.getRow(1).font = { bold: true };

    for (const row of data) {
      sheet.addRow(row);
    }
  }

  async exportExcelCat9And12(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    let data: any[];
    if (factory.trim().toUpperCase() === 'ALL') {
      const results = await Promise.all(
        FACTORY_LIST.map((f) =>
          this.getCMSByFactoryCat9(f, dateFrom, dateTo, dockeyCMS),
        ),
      );
      data = results.flat();
    } else {
      data = await this.getCMSByFactoryCat9(
        factory as FactoryCode,
        dateFrom,
        dateTo,
        dockeyCMS,
      );
    }

    sheet.columns = [
      { header: 'System', key: 'System', width: 15 },
      { header: 'Corporation', key: 'Corporation', width: 15 },
      { header: 'Factory', key: 'Factory', width: 15 },
      { header: 'Department', key: 'Department', width: 15 },
      { header: 'DocKey', key: 'DocKey', width: 15 },
      { header: 'SPeriodData', key: 'SPeriodData', width: 15 },
      { header: 'EPeriodData', key: 'EPeriodData', width: 15 },
      { header: 'ActivityType', key: 'ActivityType', width: 15 },
      { header: 'DataType', key: 'DataType', width: 15 },
      { header: 'DocType', key: 'DocType', width: 15 },
      { header: 'UndDoc', key: 'UndDoc', width: 15 },
      { header: 'DocFlow', key: 'DocFlow', width: 15 },
      { header: 'DocDate', key: 'DocDate', width: 15 },
      { header: 'DocDate2', key: 'DocDate2', width: 15 },
      { header: 'DocNo', key: 'DocNo', width: 15 },
      { header: 'UndDocNo', key: 'UndDocNo', width: 15 },
      { header: 'CustVenName', key: 'CustVenName', width: 25 },
      { header: 'InvoiceNo', key: 'InvoiceNo', width: 15 },
      { header: 'TransType', key: 'TransType', width: 15 },
      { header: 'Departure', key: 'Departure', width: 25 },
      { header: 'Destination', key: 'Destination', width: 15 },
      { header: 'PortType', key: 'PortType', width: 15 },
      { header: 'StPort', key: 'StPort', width: 15 },
      { header: 'ThPort', key: 'ThPort', width: 15 },
      { header: 'EndPort', key: 'EndPort', width: 15 },
      { header: 'Product', key: 'Product', width: 20 },
      { header: 'Quity', key: 'Quity', width: 12 },
      { header: 'Amount', key: 'Amount', width: 12 },
      { header: 'ActivityData', key: 'ActivityData', width: 15 },
      { header: 'ActivityUnit', key: 'ActivityUnit', width: 15 },
      { header: 'Unit', key: 'Unit', width: 10 },
      { header: 'UnitWeight', key: 'UnitWeight', width: 12 },
      { header: 'Memo', key: 'Memo', width: 20 },
      { header: 'CreateDateTime', key: 'CreateDateTime', width: 20 },
      { header: 'Creator', key: 'Creator', width: 15 },
      { header: 'ActivitySource', key: 'ActivitySource', width: 20 },
    ];

    sheet.getRow(1).font = { bold: true };

    for (const row of data) {
      sheet.addRow(row);
    }
  }
}
