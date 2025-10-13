import { Module } from '@nestjs/common';
import { Cat7Service } from './cat7.service';
import { Cat7Controller } from './cat7.controller';
import { DatabaseModule } from 'src/database/database.module';

@Module({
  imports: [DatabaseModule],
  controllers: [Cat7Controller],
  providers: [Cat7Service],
})
export class Cat7Module {}
