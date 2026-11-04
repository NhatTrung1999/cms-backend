import { Sequelize } from 'sequelize-typescript';
import { ConfigService } from '@nestjs/config';

/**
 * Builds a Nest provider that connects to one SQL Server database.
 * The credentials are read from the env vars `<TOKEN>_USERNAME`, `_PASSWORD`,
 * `_DATABASE_NAME`, `_HOST` and `_PORT`, so the env prefix equals the token.
 */
const createDatabaseProvider = (
  token: string,
  options: { logging?: true; pool?: { max: number; acquire: number } } = {},
) => ({
  provide: token,
  inject: [ConfigService],
  useFactory: async (configService: ConfigService) => {
    return new Sequelize({
      dialect: configService.get('DB_DIALECT'),
      username: configService.get(`${token}_USERNAME`),
      password: configService.get(`${token}_PASSWORD`),
      database: configService.get(`${token}_DATABASE_NAME`),
      host: configService.get(`${token}_HOST`),
      port: configService.get(`${token}_PORT`),
      dialectOptions: {
        options: {
          encrypt: false,
          trustServerCertificate: true,
          requestTimeout: 86400000,
        },
      },
      ...options,
    });
  },
});

const WITH_LOGGING = { logging: true } as const;

export const databaseProviders = [
  // EIP is shared by every module (ports, logs, file records...): more connections than the Sequelize
  // default (5) and a longer wait for a free one, so queued queries do not time out
  createDatabaseProvider('EIP', {
    ...WITH_LOGGING,
    pool: { max: 20, acquire: 180000 },
  }),
  // cat 9 & 12 (and cat 1 & 4)
  ...['LYV', 'LHG', 'LVL', 'LYM', 'LYF', 'JAZ', 'JZS'].map((factory) =>
    createDatabaseProvider(`${factory}_ERP`, WITH_LOGGING),
  ),
  // cat 5
  ...['LYV', 'LHG', 'LYM', 'LVL', 'JAZ', 'JZS'].map((factory) =>
    createDatabaseProvider(`${factory}_WMS`, WITH_LOGGING),
  ),
  // cat 7
  ...['LYV', 'LHG', 'LVL', 'LYM', 'JAZ', 'JZS'].map((factory) =>
    createDatabaseProvider(`${factory}_HRIS`, WITH_LOGGING),
  ),
  // cat 6 (logging is left unset on purpose, Sequelize keeps its default)
  createDatabaseProvider('UOF'),
];
