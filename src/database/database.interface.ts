export type Dialect =
  | 'mysql'
  | 'postgres'
  | 'sqlite'
  | 'mariadb'
  | 'mssql'
  | 'db2'
  | 'snowflake'
  | 'oracle';

export interface IDatabase {
  dialect?: Dialect;
  database?: string;
  username?: string;
  password?: string;
  host?: string;
  port?: number;
  logging?: boolean | ((sql: string, timing?: number) => void);
  dialectOptions?: {
    options: {
      encrypt: boolean;
      instanceName?: string;
      requestTimeout?: number;
      trustServerCertificate?: boolean;
    };
  };
}

export interface IDatabaseConfig {
  eip: IDatabase;
  erp: IDatabase;
}
