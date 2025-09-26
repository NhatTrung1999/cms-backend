import { Sequelize } from 'sequelize-typescript';
import { ConfigService } from '@nestjs/config';

export const databaseProviders = [
  {
    provide: 'EIP',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('EIP_USERNAME'),
        password: configService.get('EIP_PASSWORD'),
        database: configService.get('EIP_DATABASE_NAME'),
        host: configService.get('EIP_HOST'),
        port: configService.get('EIP_PORT'),
        dialectOptions: {
          options: {
            encrypt: false,
            trustServerCertificate: true,
          },
        },
      });
      return sequelize;
    },
  },
  {
    provide: 'ERP',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('ERP_USERNAME'),
        password: configService.get('ERP_PASSWORD'),
        database: configService.get('ERP_DATABASE_NAME'),
        host: configService.get('ERP_HOST'),
        port: configService.get('ERP_PORT'),
        dialectOptions: {
          options: {
            encrypt: false,
            trustServerCertificate: true,
          },
        },
      });
      return sequelize;
    },
  },
  {
    provide: 'LYV_WMS',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('LYV_WMS_USERNAME'),
        password: configService.get('LYV_WMS_PASSWORD'),
        database: configService.get('LYV_WMS_DATABASE_NAME'),
        host: configService.get('LYV_WMS_HOST'),
        port: configService.get('LYV_WMS_PORT'),
        dialectOptions: {
          options: {
            encrypt: false,
            trustServerCertificate: true,
          },
        },
      });
      return sequelize;
    },
  },
  {
    provide: 'LHG_WMS',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('LHG_WMS_USERNAME'),
        password: configService.get('LHG_WMS_PASSWORD'),
        database: configService.get('LHG_WMS_DATABASE_NAME'),
        host: configService.get('LHG_WMS_HOST'),
        port: configService.get('LHG_WMS_PORT'),
        dialectOptions: {
          options: {
            encrypt: false,
            trustServerCertificate: true,
          },
        },
      });
      return sequelize;
    },
  },
  {
    provide: 'LYM_WMS',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('LYM_WMS_USERNAME'),
        password: configService.get('LYM_WMS_PASSWORD'),
        database: configService.get('LYM_WMS_DATABASE_NAME'),
        host: configService.get('LYM_WMS_HOST'),
        port: configService.get('LYM_WMS_PORT'),
        dialectOptions: {
          options: {
            encrypt: false,
            trustServerCertificate: true,
          },
        },
      });
      return sequelize;
    },
  },
  {
    provide: 'LVL_WMS',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('LVL_WMS_USERNAME'),
        password: configService.get('LVL_WMS_PASSWORD'),
        database: configService.get('LVL_WMS_DATABASE_NAME'),
        host: configService.get('LVL_WMS_HOST'),
        port: configService.get('LVL_WMS_PORT'),
        dialectOptions: {
          options: {
            encrypt: false,
            trustServerCertificate: true,
          },
        },
      });
      return sequelize;
    },
  },
  {
    provide: 'JAZ_WMS',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('JAZ_WMS_USERNAME'),
        password: configService.get('JAZ_WMS_PASSWORD'),
        database: configService.get('JAZ_WMS_DATABASE_NAME'),
        host: configService.get('JAZ_WMS_HOST'),
        port: configService.get('JAZ_WMS_PORT'),
        dialectOptions: {
          options: {
            encrypt: false,
            trustServerCertificate: true,
          },
        },
      });
      return sequelize;
    },
  },
  {
    provide: 'JZS_WMS',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('JZS_WMS_USERNAME'),
        password: configService.get('JZS_WMS_PASSWORD'),
        database: configService.get('JZS_WMS_DATABASE_NAME'),
        host: configService.get('JZS_WMS_HOST'),
        port: configService.get('JZS_WMS_PORT'),
        dialectOptions: {
          options: {
            encrypt: false,
            trustServerCertificate: true,
          },
        },
      });
      return sequelize;
    },
  },
];
