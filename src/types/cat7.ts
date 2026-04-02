import { Sequelize } from 'sequelize-typescript';

export const ACTIVITY_TYPES = ['3.3'] as const;
export type ActivityType = (typeof ACTIVITY_TYPES)[number];

export type BuildQueryFn = (
  dateFrom: string,
  dateTo: string,
  eip: Sequelize,
) => Promise<string>;
