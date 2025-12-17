import { Module } from '@nestjs/common';
import { Cat1andcat4Service } from './cat1andcat4.service';
import { Cat1andcat4Controller } from './cat1andcat4.controller';
import { DatabaseModule } from 'src/database/database.module';
import { MulterModule } from '@nestjs/platform-express';
import { diskStorage } from 'multer';
import { extname } from 'path';

@Module({
  imports: [
    DatabaseModule,
    MulterModule.register({
      storage: diskStorage({
        destination: './uploads',
        filename: (req, file, cb) => {
          const uniqueSuffix =
            Date.now() + '-' + Math.round(Math.random() * 1e9);
          const ext = extname(file.originalname);
          cb(null, `${file.fieldname}-${uniqueSuffix}${ext}`);
        },
      }),
      fileFilter: (req, file, cb) => {
        if (!file.originalname.match(/\.(xlsx|xls)$/)) {
          return cb(new Error('Only Excel files are allowed!'), false);
        }
        cb(null, true);
      },
    }),
  ],
  controllers: [Cat1andcat4Controller],
  providers: [Cat1andcat4Service],
})
export class Cat1andcat4Module {}
