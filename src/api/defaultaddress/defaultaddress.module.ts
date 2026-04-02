import { Module } from '@nestjs/common';
import { DefaultaddressService } from './defaultaddress.service';
import { DefaultaddressController } from './defaultaddress.controller';
import { DatabaseModule } from 'src/database/database.module';

@Module({
  imports: [DatabaseModule],
  controllers: [DefaultaddressController],
  providers: [DefaultaddressService],
})
export class DefaultaddressModule {}
