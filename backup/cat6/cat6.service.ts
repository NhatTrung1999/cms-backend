// import { Inject, Injectable } from '@nestjs/common';
// import { Sequelize } from 'sequelize-typescript';
// import { QueryTypes } from 'sequelize';
// import { ICat6Data, ICat6Query, ICat6Record } from 'src/types/cat6';
// import dayjs from 'dayjs';
// import {
//   AutoCMSCat6Data,
//   fakeCat6LHGData,
//   fakeCat6LVLData,
//   fakeCat6LYMData,
//   fakeCat6LYVData,
// } from 'src/fakedata';
// import { getFactory } from 'src/helper/factory.helper';
// import {
//   normalizeRoute,
//   splitRecordByAirport,
//   sumHotelNights,
// } from 'src/helper/cat6.helper';
// dayjs().format();

// @Injectable()
// export class Cat6Service {
//   constructor(@Inject('UOF') private readonly UOF: Sequelize) {}

//   async getDataCat6(
//     dateFrom: string,
//     dateTo: string,
//     factory: string,
//     page: number = 1,
//     limit: number = 20,
//     sortField: string = 'Document_Date',
//     sortOrder: string = 'asc',
//   ) {
//     const offset = (page - 1) * limit;
//     let where = `WHERE 1=1 AND BPMStatus = 'F'
//                   AND ISNULL(Accommodation ,'')<>''
//                   AND ISNULL(Factory_User ,'')<>''
//                   AND DOC_NBR = 'LYV-HR-BT251200011'`;
//     const replacements: any[] = [];

//     if (dateTo && dateFrom) {
//       where += ` AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?`;
//       replacements.push(dateFrom, dateTo);
//     }

//     if (factory) {
//       where += ` AND Factory_User LIKE ?`;
//       replacements.push(`%${factory}%`);
//     }

//     const query = `SELECT *, COUNT(TripID) OVER() AS TotalRow
//                     FROM CDS_HRBUSS_BusTripData
//                     ${where}`;

//     const dataResults = (await this.UOF.query(query, {
//       type: QueryTypes.SELECT,
//       replacements,
//     })) as ICat6Query[];

//     const records: ICat6Record[] = dataResults
//       .map((item) => ({
//         ...item,
//         Routes: item.Routes
//           ? JSON.parse(item.Routes).map((r: any) => normalizeRoute(r))
//           : [],
//         Accommodation: item.Accommodation ? JSON.parse(item.Accommodation) : [],
//       }))
//       .flatMap((record) => splitRecordByAirport(record));

//     return records;
//     // let data = records.map((item) => {
//     //   const routes = item.Routes || [];
//     //   const flightSegment = routes.find((r) => r.isAirport);
//     //   const startSegment = routes[0];

//     //   const otherDestinations = routes.filter(
//     //     (r) => r !== startSegment && !r.isAirport,
//     //   );

//     //   const getAddr = (r: any) => r?.AddressDetail || r?.AddressName || '';
//     //   const totalNights = sumHotelNights(item.Accommodation);

//     //   let businessTripType: string = '';
//     //   switch (item.Factory.trim().toLowerCase()) {
//     //     case 'factory_domestic':
//     //       businessTripType = 'Domestic business trip within the group';
//     //       break;
//     //     case 'outside_domestic':
//     //       businessTripType = 'Domestic business trip to third-party entities';
//     //       break;
//     //     case 'factory_oversea':
//     //       businessTripType = 'Overseas business trip within the group';
//     //       break;
//     //     case 'outside_oversea':
//     //       businessTripType = 'Overseas business trip to third-party entities';
//     //       break;
//     //     default:
//     //       businessTripType = '';
//     //       break;
//     //   }

//     //   return {
//     //     Document_Date: item.CreatedAt,
//     //     Document_Number: item.DOC_NBR,
//     //     Staff_ID: item.UserCreate,
//     //     Round_trip_One_way:
//     //       item.TypeTravel.trim().toLowerCase() === 'round'.trim().toLowerCase()
//     //         ? 'Round trip'
//     //         : 'One-way',
//     //     Start_Time: item.DateStart,
//     //     End_Time: item.DateEnd,
//     //     Business_Trip_Type: businessTripType,
//     //     Place_of_Departure: getAddr(startSegment),
//     //     Land_Trasportation_Type_A: startSegment?.Transport || '',
//     //     Land_Transport_Distance_km_A: 'API Calculation',
//     //     Departure_Airport: flightSegment?.From || '',
//     //     Destination_Airport: flightSegment?.To || '',
//     //     Air_Transport_Distance_km: flightSegment ? 'API Calculation' : '',
//     //     Third_country_transfer_Destination: getAddr(otherDestinations[0]),
//     //     Land_Transportation_Type_B: otherDestinations[0]?.Transport || '',
//     //     Land_Transport_Distance_km_B: otherDestinations[0]
//     //       ? 'API Calculation'
//     //       : '',
//     //     Destination_2: getAddr(otherDestinations[1]),
//     //     Destination_3: getAddr(otherDestinations[2]),
//     //     Destination_4: getAddr(otherDestinations[3]),
//     //     Destination_5: getAddr(otherDestinations[4]),
//     //     Destination_6: getAddr(otherDestinations[5]),
//     //     Land_Transportation_Type: otherDestinations[0]?.Transport || '',
//     //     Land_Transport_Distance_km: 'API Calculation',
//     //     Number_of_nights_stayed: totalNights,
//     //     TotalRow: item.TotalRow || 0,
//     //   };
//     // });

//     // data.sort((a, b) => {
//     //   const aValue = a[sortField];
//     //   const bValue = b[sortField];
//     //   if (sortOrder === 'asc') {
//     //     return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
//     //   } else {
//     //     return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
//     //   }
//     // });

//     // data = data.slice(offset, offset + limit);

//     // const total = data[0]?.TotalRow || 0;

//     // const hasMore = offset + data.length < total;

//     // return { data, page, limit, total, hasMore };
//   }

//   async transformData(
//     data: AutoCMSCat6Data[],
//     factory: string,
//     dateFrom: string,
//     dateTo: string,
//   ) {
//     return data.flatMap((item) => {
//       const routes: { transType: string; dep: string; dest: string }[] = [];

//       if (item.PlaceOfDeparture && item.DepartureAirport) {
//         routes.push({
//           transType: item.LandTrasportationTypeA,
//           dep: item.PlaceOfDeparture,
//           dest: item.DepartureAirport,
//         });
//       }

//       if (item.DepartureAirport && item.DestinationAirport) {
//         routes.push({
//           transType: 'Flight',
//           dep: item.DepartureAirport,
//           dest: item.DestinationAirport,
//         });
//       }

//       if (item.DestinationAirport && item.ThirdCountryTransferDestination) {
//         routes.push({
//           transType: item.LandTrasportationTypeB,
//           dep: item.DestinationAirport,
//           dest: item.ThirdCountryTransferDestination,
//         });
//       }

//       return routes.map((route, index) => {
//         let transType = '';
//         switch (route.transType.trim().toLowerCase()) {
//           case 'Car'.trim().toLowerCase():
//             transType = '計程車/出租車';
//             break;
//           case 'Company Shuttle Car'.trim().toLowerCase():
//             transType = '公司車';
//             break;
//           case 'Flight'.trim().toLowerCase():
//             transType = '飛機';
//             break;
//           default:
//             transType = route.transType.trim();
//             break;
//         }

//         let docKey = '';
//         switch (item.BusinessTripType.trim().toLowerCase()) {
//           case 'Domestic business trip within the group'.trim().toLowerCase():
//             docKey = '3.5.1';
//             break;
//           case 'Domestic business trip to third party entities'
//             .trim()
//             .toLowerCase():
//             docKey = '3.5.2';
//             break;
//           case 'Overseas business trip within the group'.trim().toLowerCase():
//             docKey = '3.5.3';
//             break;
//           case 'Overseas business trip to third party entities'
//             .trim()
//             .toLowerCase():
//             docKey = '3.5.4';
//             break;
//           default:
//             break;
//         }
//         return {
//           System: 'BPM',
//           Corporation: 'LAI YIH',
//           Factory: getFactory(factory),
//           Department: item.Dept,
//           DocKey: docKey,
//           ActivitySource: '',
//           SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
//           EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
//           ActivityType: '3.5',
//           DataType: '2',
//           DocType: '洽公單',
//           DocDate: dayjs(item.DocumentDate).format('YYYY/MM/DD'),
//           DocDate2: dayjs(item.StartTime).format('YYYY/MM/DD'),
//           DocNo: item.DocumentNumber,
//           UndDocNo: `${item.DocumentNumber}-${index + 1}`,
//           TransType: transType,
//           Departure: route.dep,
//           Destination: route.dest,
//           Memo: item.BusinessTripType,
//           CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
//           Creator: '',
//         };
//       });
//     });
//   }

//   async autoSentCMS(dateFrom: string, dateTo: string, factory: string) {
//     return await this.transformData(fakeCat6LVLData, factory, dateFrom, dateTo);
//   }

//   async autoSentCMSV2(dateFrom: string, dateTo: string, factory: string) {
//     return fakeCat6LVLData.map((item) => {
//       let docKey = '';
//       switch (item.BusinessTripType.trim().toLowerCase()) {
//         case 'Domestic business trip within the group'.trim().toLowerCase():
//           docKey = '3.5.1';
//           break;
//         case 'Domestic business trip to third party entities'
//           .trim()
//           .toLowerCase():
//           docKey = '3.5.2';
//           break;
//         case 'Overseas business trip within the group'.trim().toLowerCase():
//           docKey = '3.5.3';
//           break;
//         case 'Overseas business trip to third party entities'
//           .trim()
//           .toLowerCase():
//           docKey = '3.5.4';
//           break;
//         default:
//           break;
//       }
//       return {
//         System: 'BPM',
//         Corporation: 'LAI YIH',
//         Factory: getFactory(factory),
//         Department: item.Dept,
//         DocKey: docKey,
//         ActivitySource: '住宿',
//         SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
//         EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
//         ActivityType: '3.5',
//         DataType: '3',
//         DocType: '出差住宿單',
//         DocDate: dayjs(item.DocumentDate).format('YYYY/MM/DD'),
//         DocDate2: dayjs(item.StartTime).format('YYYY/MM/DD'),
//         DocNo: item.DocumentNumber,
//         UndDocNo: item.DocumentNumber,
//         TransType: 'double',
//         ActivityData: item.NumberOfNightsStayed,
//         Memo: '',
//         CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
//         Creator: '',
//       };
//     });
//   }
// }
