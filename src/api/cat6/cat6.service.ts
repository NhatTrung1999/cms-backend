import { Inject, Injectable } from '@nestjs/common';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';
import { ICat6Data, ICat6Query, ICat6Record } from 'src/types/cat6';
import dayjs from 'dayjs';
import { fakeCat6Data } from 'src/fakedata';
import { getFactory } from 'src/helper/factory.helper';
import { normalizeRoute, splitRecordByAirport } from 'src/helper/cat6.helper';
dayjs().format();

@Injectable()
export class Cat6Service {
  constructor(@Inject('UOF') private readonly UOF: Sequelize) {}

  async getDataCat6(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'Document_Date',
    sortOrder: string = 'asc',
  ) {
    const offset = (page - 1) * limit;
    let where = 'WHERE 1=1';
    const replacements: any[] = [];

    if (dateTo && dateFrom) {
      where += ` AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }

    if (factory) {
      where += ` AND Factory_User LIKE ?`;
      replacements.push(`%${factory}%`);
    }

    const query = `SELECT *, COUNT(TripID) OVER() AS TotalRow
                    FROM CDS_HRBUSS_BusTripData
                    ${where}`;

    // const countQuery = `SELECT COUNT(*) as total
    //                     FROM CDS_HRBUSS_BusTripData
    //                     WHERE DOC_NBR = 'LYV_HR_BT260100010'`;

    // const [dataResults, countResults] = await Promise.all([
    //   this.UOF.query(query, { type: QueryTypes.SELECT }) as Promise<
    //     ICat6Query[]
    //   >,
    //   this.UOF.query(countQuery, {
    //     type: QueryTypes.SELECT,
    //   }) as Promise<{ total: number }[]>,
    // ]);

    const dataResults = (await this.UOF.query(query, {
      type: QueryTypes.SELECT,
      replacements,
    })) as ICat6Query[];

    const records: ICat6Record[] = dataResults
      .map((item) => ({
        ...item,
        Routes: item.Routes
          ? JSON.parse(item.Routes).map((r: any) => normalizeRoute(r))
          : [],
        Accommodation: item.Accommodation ? JSON.parse(item.Accommodation) : [],
      }))
      .flatMap((record) => splitRecordByAirport(record));

    let data = records.map((item) => {
      const routes = item.Routes || [];
      const flightSegment = routes.find((r) => r.isAirport);
      const startSegment = routes[0];

      const otherDestinations = routes.filter(
        (r) => r !== startSegment && !r.isAirport,
      );

      const getAddr = (r: any) => r?.AddressDetail || r?.AddressName || '';
      const totalNights =
        item.Accommodation?.filter(
          (item) => item.isSameAsAbove === false,
        ).reduce((sum, acc) => sum + (acc.nights || 0), 0) || 0;

      let businessTripType: string = '';
      switch (item.Factory.trim().toLowerCase()) {
        case 'factory_domestic':
          businessTripType = 'Domestic business trip within the group';
          break;
        case 'outside_domestic':
          businessTripType = 'Domestic business trip to third-party entities';
          break;
        case 'factory_oversea':
          businessTripType = 'Overseas business trip within the group';
          break;
        case 'outside_oversea':
          businessTripType = 'Overseas business trip to third-party entities';
          break;
        default:
          businessTripType = '';
          break;
      }

      return {
        Document_Date: item.CreatedAt,
        Document_Number: item.DOC_NBR,
        Staff_ID: item.UserCreate,
        Round_trip_One_way:
          item.TypeTravel.trim().toLowerCase() === 'round'.trim().toLowerCase()
            ? 'Round trip'
            : 'One-way',
        Start_Time: item.DateStart,
        End_Time: item.DateEnd,
        Business_Trip_Type: businessTripType,
        Place_of_Departure: getAddr(startSegment),
        Land_Trasportation_Type_A: startSegment?.Transport || '',
        Land_Transport_Distance_km_A: 'API Calculation',
        Departure_Airport: flightSegment?.From || '',
        Destination_Airport: flightSegment?.To || '',
        Air_Transport_Distance_km: flightSegment ? 'API Calculation' : '',
        Third_country_transfer_Destination: getAddr(otherDestinations[0]),
        Land_Transportation_Type_B: otherDestinations[0]?.Transport || '',
        Land_Transport_Distance_km_B: otherDestinations[0]
          ? 'API Calculation'
          : '',
        Destination_2: getAddr(otherDestinations[1]),
        Destination_3: getAddr(otherDestinations[2]),
        Destination_4: getAddr(otherDestinations[3]),
        Destination_5: getAddr(otherDestinations[4]),
        Destination_6: getAddr(otherDestinations[5]),
        Land_Transportation_Type: otherDestinations[0]?.Transport || '',
        Land_Transport_Distance_km: 'API Calculation',
        Number_of_nights_stayed: totalNights,
        TotalRow: item.TotalRow || 0,
      };
    });

    // console.log(formattedData);
    // const data = records.map((item) => {
    //   return {
    //     Document_Date: item.CreatedAt,
    //     Document_Number: item.DOC_NBR,
    //     Staff_ID: item.UserCreate,
    //     Round_trip_One_way: item.TypeTravel,
    //     Start_Time: item.DateStart,
    //     End_Time: item.DateEnd,
    //     Business_Trip_Type: item.Factory,
    //     Place_of_Departure: item.Routes[0].AddressDetail,
    //     Departure_Airport: item.Routes[1].From,
    //     Land_Transport_Distance_km_A: 'API Calculation',
    //     Land_Trasportation_Type_A: item.Routes[0].Transport,
    //     Destination_Airport: item.Routes[1].To,
    //     Third_country_transfer_Destination: item.Routes[2].AddressDetail,
    //     Land_Transport_Distance_km_B: 'API Calculation',
    //     Land_Transportation_Type_B: item.Routes[2].Transport,
    //     Destination_2: item.Routes[3].AddressDetail,
    //     Destination_3: '',
    //     Destination_4: '',
    //     Destination_5: '',
    //     Destination_6: '',
    //     Land_Transport_Distance_km: 'API Calculation',
    //     Land_Transportation_Type: item.Routes[2].Transport,
    //     Air_Transport_Distance_km: 'API Calculation',
    //     Number_of_nights_stayed: 0,
    //   };
    // });

    // return records;

    // let data: ICat6Data[] = records.map((item) => {
    //   const routes = item.Routes.filter((item) => item.isAirport === true);
    // let businessTripType: string = '';
    // switch (item.Factory) {
    //   case 'factory_domestic':
    //     businessTripType = 'Domestic business trip within the group';
    //     break;
    //   case 'outside_domestic':
    //     businessTripType = 'Domestic business trip to third-party entities';
    //     break;
    //   case 'factory_oversea':
    //     businessTripType = 'Overseas business trip within the group';
    //     break;
    //   case 'outside_oversea':
    //     businessTripType = 'Overseas business trip to third-party entities';
    //     break;
    //   default:
    //     businessTripType = '';
    //     break;
    // }

    //   return {
    //     Document_Date: item.CreatedAt,
    //     Document_Number: item.DOC_NBR,
    //     Staff_ID: item.UserCreate,
    //     Round_trip_One_way:
    //       item.TypeTravel === 'round' ? 'Round trip' : 'One-way',
    //     Business_Trip_Type: businessTripType,
    //     Place_of_Departure: routes[0].AddressDetail,
    //     Departure_Airport: routes[0].AddressName,
    //     Land_Transport_Distance_km_A: 0,
    //     Land_Trasportation_Type_A: 0,
    //     Destination_Airport: routes[routes.length - 1].AddressName,
    //     Destination_1: '',
    //     Destination_2: '',
    //     Land_Transport_Distance_km_B: 0,
    //     Land_Trasportation_Type_B: 0,
    //     Air_Transport_Distance_km: 0,
    //     Number_of_nights_stayed: item.StayNight,
    //   };
    // });

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

    const total = data[0]?.TotalRow || 0;

    const hasMore = offset + data.length < total;

    return { data, page, limit, total, hasMore };
  }

  async autoSentCMS(dateFrom: string, dateTo: string, factory: string) {
    // return fakeCat6Data.map((item) => ({
    // System: 'BPM',
    // Corporation: getFactory(factory),
    // Factory: getFactory(factory),
    // Department: '設計部',
    // DocKey: item.DocumentNumber,
    // ActivitySource: '',
    // SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
    // EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
    // ActivityType: '3.5',
    // DataType: '2',
    // DocType: item.DocumentNumber,
    // DocDate: dayjs(item.DocumentDate).format('YYYY/MM/DD'),
    // DocDate2: dayjs(item.StartTime).format('YYYY/MM/DD'),
    // DocNo: item.DocumentNumber,
    // UndDocNo: item.DocumentNumber,
    // TransType: item.LandTrasportationTypeA,
    // Departure: item.PlaceOfDeparture,
    // Destination: item.ThirdCountryTransferDestination,
    // Memo: item.BusinessTripType,
    // CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
    // Creator: '',
    // }));
    return [
      {
        System: 'BPM',
        Corporation: getFactory(factory),
        Factory: getFactory(factory),
        Department: '設計部',
        DocKey: '3.5.4',
        ActivitySource: '',
        SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
        EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
        ActivityType: '3.5',
        DataType: '2',
        DocType: 'LYV-HR-BT250100001',
        DocDate: dayjs('2025-01-02').format('YYYY/MM/DD'),
        DocDate2: dayjs('2025-02-15').format('YYYY/MM/DD'),
        DocNo: 'LYV-HR-BT250100001',
        UndDocNo: 'LYV-HR-BT250100001-1',
        TransType: 'Company Shuttle Bus',
        Departure:
          '3-5 Đ. Tên Lửa, An Lạc, Bình Tân, Thành phố Hồ Chí Minh 763430越南',
        Destination: 'SGN',
        Memo: 'Overseas business trip to third-party entities',
        CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
        Creator: '',
      },
      {
        System: 'BPM',
        Corporation: getFactory(factory),
        Factory: getFactory(factory),
        Department: '設計部',
        DocKey: '3.5.4',
        ActivitySource: '',
        SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
        EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
        ActivityType: '3.5',
        DataType: '2',
        DocType: 'LYV-HR-BT250100001',
        DocDate: dayjs('2025-01-02').format('YYYY/MM/DD'),
        DocDate2: dayjs('2025-02-15').format('YYYY/MM/DD'),
        DocNo: 'LYV-HR-BT250100001',
        UndDocNo: 'LYV-HR-BT250100001-2',
        TransType: '',
        Departure: 'SGN',
        Destination: 'NUE',
        Memo: 'Overseas business trip to third-party entities',
        CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
        Creator: '',
      },
      {
        System: 'BPM',
        Corporation: getFactory(factory),
        Factory: getFactory(factory),
        Department: '設計部',
        DocKey: '3.5.4',
        ActivitySource: '',
        SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
        EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
        ActivityType: '3.5',
        DataType: '2',
        DocType: 'LYV-HR-BT250100001',
        DocDate: dayjs('2025-01-02').format('YYYY/MM/DD'),
        DocDate2: dayjs('2025-02-15').format('YYYY/MM/DD'),
        DocNo: 'LYV-HR-BT250100001',
        UndDocNo: 'LYV-HR-BT250100001-3',
        TransType: 'Car',
        Departure: 'NUE',
        Destination: 'Adi-Dassler-Straße 1, 91074 Herzogenaurach, Germany',
        Memo: 'Overseas business trip to third-party entities',
        CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
        Creator: '',
      },
    ];
    // return [
    //   {
    // System: 'BPM',
    // Corporation: '樂億 - LYV',
    // Factory: '樂億 - LYV',
    // Department: '設計部',
    // DocKey: 'LYV-HR-BT250100001',
    // ActivitySource: '',
    // SPeriodData: '2026/01/24',
    // EPeriodData: '2026/01/24',
    // ActivityType: '3.5',
    // DataType: '2',
    // DocType: 'LYV-HR-BT250100001',
    // DocDate: '2025/01/02',
    // DocDate2: '2025/02/15',
    // DocNo: 'LYV-HR-BT250100001',
    // UndDocNo: 'LYV-HR-BT250100001-1',
    // TransType: 'Company Shuttle Bus',
    // Departure:
    //   '3-5 Đ. Tên Lửa, An Lạc, Bình Tân, Thành phố Hồ Chí Minh 763430越南',
    // Destination: 'SGN',
    // Memo: 'Round Trip',
    // CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
    // Creator: '',
    //   },
    // ];
  }

  async autoSentCMSV2(dateFrom: string, dateTo: string, factory: string) {
    return fakeCat6Data.map((item) => ({
      System: 'BPM',
      Corporation: getFactory(factory),
      Factory: getFactory(factory),
      Department: '設計部',
      DocKey: item.StaffID.toString(),
      ActivitySource: '住宿',
      SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
      EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
      ActivityType: '3.5',
      DataType: '3',
      DocType: '出差住宿單',
      DocDate: dayjs(item.DocumentDate).format('YYYY/MM/DD'),
      DocDate2: dayjs(item.StartTime).format('YYYY/MM/DD'),
      DocNo: item.DocumentNumber,
      UndDocNo: item.DocumentNumber,
      TransType: 'double',
      ActivityData: item.NumberOfNightsStayed.toString(),
      Memo: '',
      CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
      Creator: '',
    }));
  }
}
