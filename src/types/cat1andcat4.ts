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

export interface IDataStyleAutoFill {
  Id: string;
  PrefixOfMatCode: string;
  Style: string;
  CreatedBy: string;
  CreatedFactory: string;
  CreatedAt: string;
  UpdatedBy: string;
  UpdatedFactory: string;
  UpdatedAt: string;
}

export const ACTIVITY_TYPES = ['3.1', '4.1'] as const;
export type ActivityType = (typeof ACTIVITY_TYPES)[number];
