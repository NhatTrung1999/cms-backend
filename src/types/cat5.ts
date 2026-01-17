export interface IDataCat5 {
  Waste_disposal_date: string;
  Vender_Name: string;
  Waste_collection_address: string;
  Transportation_Distance_km: string;
  The_type_of_waste: string;
  Waste_type: string;
  Waste_Treatment_method: string;
  Weight_of_waste_treated_Unit_kg: string;
  TKT_Ton_km: string;
}


// export const ACTIVITY_TYPES = ['3.6', '4.4'] as const
export const ACTIVITY_TYPES = ['4.4'] as const;
export type ActivityType = typeof ACTIVITY_TYPES[number]