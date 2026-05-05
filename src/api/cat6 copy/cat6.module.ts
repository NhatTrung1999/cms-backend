import { Module } from '@nestjs/common';
import { Cat6Service } from './cat6.service';
import { Cat6Controller } from './cat6.controller';
import { DatabaseModule } from 'src/database/database.module';

@Module({
  imports: [DatabaseModule],
  controllers: [Cat6Controller],
  providers: [Cat6Service],
})
export class Cat6Module {}
