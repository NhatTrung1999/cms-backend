export interface ICat6Query {
  TripID: number;
  DOC_NBR: string;
  Factory: string;
  Departure: string;
  Destination: string;
  TypeTravel: string;
  Routes: string;
  DuringDay: number;
  Distance: number;
  TotalMoney: number;
  Accommodation: string;
  CreatedAt: string;
  UserCreate: string;
  BPMStatus: string;
  YN: number;
  Factory_User: string;
  DateStart: string;
  DateEnd: string;
}

export interface ICat6Record
  extends Omit<ICat6Query, 'Routes' | 'Accommodation'> {
  Routes: {
    AddressName: string;
    Transport: string;
    AddressDetail: string;
    isAirport: boolean;
  }[];

  Accommodation: {
    id: string;
    isSameAsAbove: boolean;
    type: string;
    address: string;
    nights: number;
  }[];
}


export interface ICat6Data {
  Document_Date: string;
  Document_Number: string;
  Staff_ID: string;
  Round_trip_One_way: string;
  Business_Trip_Type: string;
  Place_of_Departure: string;
  Departure_Airport: string;
  Land_Transport_Distance_km_A: number;
  Land_Trasportation_Type_A: number;
  Destination_Airport: string;
  Destination_1: string;
  Destination_2: string;
  Land_Transport_Distance_km_B: number;
  Land_Trasportation_Type_B: number;
  Air_Transport_Distance_km: number;
  Number_of_nights_stayed: number;
}
