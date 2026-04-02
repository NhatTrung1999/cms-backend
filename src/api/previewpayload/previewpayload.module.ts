import { Module } from '@nestjs/common';
import { PreviewpayloadService } from './previewpayload.service';
import { PreviewpayloadController } from './previewpayload.controller';
import { DatabaseModule } from 'src/database/database.module';
import { EventsModule } from '../events/events.module';

@Module({
  imports: [DatabaseModule, EventsModule],
  controllers: [PreviewpayloadController],
  providers: [PreviewpayloadService],
})
export class PreviewpayloadModule {}
