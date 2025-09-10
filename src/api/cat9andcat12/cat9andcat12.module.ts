import { Module } from '@nestjs/common';
import { Cat9andcat12Service } from './cat9andcat12.service';
import { Cat9andcat12Controller } from './cat9andcat12.controller';
import { DatabaseModule } from 'src/database/database.module';

@Module({
  imports: [DatabaseModule],
  controllers: [Cat9andcat12Controller],
  providers: [Cat9andcat12Service],
})
export class Cat9andcat12Module {}
