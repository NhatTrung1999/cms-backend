import { Module } from '@nestjs/common';
import { Cat5Service } from './cat5.service';
import { Cat5Controller } from './cat5.controller';
import { DatabaseModule } from 'src/database/database.module';

@Module({
  imports: [DatabaseModule],
  controllers: [Cat5Controller],
  providers: [Cat5Service],
})
export class Cat5Module {}
