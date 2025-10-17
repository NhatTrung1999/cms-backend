import { Module } from '@nestjs/common';
import { Cat1andcat4Service } from './cat1andcat4.service';
import { Cat1andcat4Controller } from './cat1andcat4.controller';
import { DatabaseModule } from 'src/database/database.module';

@Module({
  imports: [DatabaseModule],
  controllers: [Cat1andcat4Controller],
  providers: [Cat1andcat4Service],
})
export class Cat1andcat4Module {}
