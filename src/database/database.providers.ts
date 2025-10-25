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
    provide: 'LYV_ERP',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('LYV_ERP_USERNAME'),
        password: configService.get('LYV_ERP_PASSWORD'),
        database: configService.get('LYV_ERP_DATABASE_NAME'),
        host: configService.get('LYV_ERP_HOST'),
        port: configService.get('LYV_ERP_PORT'),
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
    provide: 'LHG_ERP',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('LHG_ERP_USERNAME'),
        password: configService.get('LHG_ERP_PASSWORD'),
        database: configService.get('LHG_ERP_DATABASE_NAME'),
        host: configService.get('LHG_ERP_HOST'),
        port: configService.get('LHG_ERP_PORT'),
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
    provide: 'LVL_ERP',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('LVL_ERP_USERNAME'),
        password: configService.get('LVL_ERP_PASSWORD'),
        database: configService.get('LVL_ERP_DATABASE_NAME'),
        host: configService.get('LVL_ERP_HOST'),
        port: configService.get('LVL_ERP_PORT'),
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
    provide: 'LYM_ERP',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('LYM_ERP_USERNAME'),
        password: configService.get('LYM_ERP_PASSWORD'),
        database: configService.get('LYM_ERP_DATABASE_NAME'),
        host: configService.get('LYM_ERP_HOST'),
        port: configService.get('LYM_ERP_PORT'),
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
    provide: 'LYF_ERP',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('LYF_ERP_USERNAME'),
        password: configService.get('LYF_ERP_PASSWORD'),
        database: configService.get('LYF_ERP_DATABASE_NAME'),
        host: configService.get('LYF_ERP_HOST'),
        port: configService.get('LYF_ERP_PORT'),
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
  {
    provide: 'HRIS',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('HRIS_USERNAME'),
        password: configService.get('HRIS_PASSWORD'),
        database: configService.get('HRIS_DATABASE_NAME'),
        host: configService.get('HRIS_HOST'),
        port: configService.get('HRIS_PORT'),
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
    provide: 'UOF',
    inject: [ConfigService],
    useFactory: async (configService: ConfigService) => {
      const sequelize = new Sequelize({
        dialect: configService.get('DB_DIALECT'),
        username: configService.get('UOF_USERNAME'),
        password: configService.get('UOF_PASSWORD'),
        database: configService.get('UOF_DATABASE_NAME'),
        host: configService.get('UOF_HOST'),
        port: configService.get('UOF_PORT'),
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
