export interface IDataCat9AndCat12 {
  No: number;
  Date: string;
  Invoice_Number: string;
  Article_Name: string;
  Quantity: number;
  Gross_Weight: number;
  Customer_ID: string;
  Local_Land_Transportation: string;
  Port_Of_Departure: string;
  Port_Of_Arrival: string;
  Land_Transport_Distance: number;
  Sea_Transport_Distance: number;
  Air_Transport_Distance: number;
  Transport_Method: string;
  Land_Transport_Ton_Kilometers: number;
  Sea_Transport_Ton_Kilometers: number;
  Air_Transport_Ton_Kilometers: number;
}

export interface IDataPortCode {
  Id: string;
  CustomerNumber: string;
  CustomerName: string;
  TWCustomerName: string;
  Country: string;
  PortCode: string;
  CreatedAt: string;
  CreatedFactory: string;
  CreatedDate: string;
  UpdatedAt: string;
  UpdatedFactory: string;
  UpdatedDate: string;
}

// export const ACTIVITY_TYPES = ['3.2', '5.3'] as const;
export const ACTIVITY_TYPES = ['5.3'] as const;
// export const ACTIVITY_TYPES = ['3.2'] as const;
export type ActivityType = (typeof ACTIVITY_TYPES)[number];
