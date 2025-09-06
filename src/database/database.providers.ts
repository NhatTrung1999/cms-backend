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
];
