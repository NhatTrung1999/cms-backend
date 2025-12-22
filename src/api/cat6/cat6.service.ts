import { Inject, Injectable } from '@nestjs/common';
import { CreateCat6Dto } from './dto/create-cat6.dto';
import { UpdateCat6Dto } from './dto/update-cat6.dto';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';
import { ICat6Data, ICat6Query, ICat6Record } from 'src/types/cat6';

interface Route {
  AddressName?: string;
  AddressDetail?: string;
  Transport: string;
  isAirport: boolean;
  From?: string;
  To?: string;
}

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
    const query = `SELECT *
                    FROM CDS_HRBUSS_BusTripData
                    WHERE DOC_NBR = 'LYV_HR_BT251200015'`;

    const countQuery = `SELECT COUNT(*) as total
                        FROM CDS_HRBUSS_BusTripData
                        WHERE DOC_NBR = 'LYV_HR_BT251200015'`;

    const [dataResults, countResults] = await Promise.all([
      this.UOF.query(query, { type: QueryTypes.SELECT }) as Promise<
        ICat6Query[]
      >,
      this.UOF.query(countQuery, {
        type: QueryTypes.SELECT,
      }) as Promise<{ total: number }[]>,
    ]);

    const records: ICat6Record[] = dataResults
      .map((item) => ({
        ...item,
        Routes: item.Routes
          ? JSON.parse(item.Routes).map((r: any) => this.normalizeRoute(r))
          : [],
        Accommodation: item.Accommodation ? JSON.parse(item.Accommodation) : [],
      }))
      .flatMap((record) => this.splitRecordByAirport(record));

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

    // console.log(records);
    return records;
    // return data

    // let data: ICat6Data[] = records.map((item) => {
    //   const routes = item.Routes.filter((item) => item.isAirport === true);
    //   let businessTripType: string = '';
    //   switch (item.Factory) {
    //     case 'factory_domestic':
    //       businessTripType = 'Domestic business trip within the group';
    //       break;
    //     case 'outside_domestic':
    //       businessTripType = 'Domestic business trip to third-party entities';
    //       break;
    //     case 'factory_oversea':
    //       businessTripType = 'Overseas business trip within the group';
    //       break;
    //     case 'outside_oversea':
    //       businessTripType = 'Overseas business trip to third-party entities';
    //       break;
    //     default:
    //       businessTripType = '';
    //       break;
    //   }

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

    // data.sort((a, b) => {
    //   const aValue = a[sortField];
    //   const bValue = b[sortField];
    //   if (sortOrder === 'asc') {
    //     return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
    //   } else {
    //     return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
    //   }
    // });

    // data = data.slice(offset, offset + limit);

    // const total = countResults[0]?.total || 0;

    // const hasMore = offset + data.length < total;

    // return { data, page, limit, total, hasMore };
  }

  private splitRecordByAirport(record: ICat6Record): ICat6Record[] {
    const routes = record.Routes;
    const result: ICat6Record[] = [];

    let startIndex = 0;
    let airportCount = 0;

    for (let i = 0; i < routes.length; i++) {
      if (routes[i].isAirport) {
        airportCount++;

        if (airportCount > 1) {
          const cutIndex = i - 1;

          result.push({
            ...record,
            Routes: routes.slice(startIndex, cutIndex + 1),
          });

          startIndex = cutIndex;
        }
      }
    }

    result.push({
      ...record,
      Routes: routes.slice(startIndex),
    });

    return result;
  }

  private normalizeRoute(route: any): Route {
    return {
      AddressName: route.AddressName ?? null,
      AddressDetail: route.AddressDetail ?? null,
      Transport: route.Transport,
      isAirport: route.isAirport,
      From: route.From ?? null,
      To: route.To ?? null,
    };
  }
}
