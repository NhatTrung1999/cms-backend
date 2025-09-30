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
  CustomerName: null;
  TWCustomerName: null;
  Country: null;
  PortCode: string;
  CreatedAt: string;
  CreatedFactory: string;
  CreatedDate: string;
}
