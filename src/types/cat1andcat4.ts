export interface IDataPortCodeCat1AndCat4 {
  Id: string;
  SupplierID: string;
  SupplierName: string;
  TWSupplierName: string;
  Country: string;
  PortCode: string;
  FactoryCode: string;
  CreatedBy: string;
  CreatedFactory: string;
  CreatedDate: string;
  UpdatedBy: string;
  UpdatedFactory: string;
  UpdatedDate: string;
}


export const ACTIVITY_TYPES = ['3.1', '4.1'] as const;
export type ActivityType = (typeof ACTIVITY_TYPES)[number];
