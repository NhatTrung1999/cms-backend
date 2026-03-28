import { Module } from '@nestjs/common';
import { PreviewPayloadService } from './preview-payload.service';
import { PreviewPayloadController } from './preview-payload.controller';

@Module({
  controllers: [PreviewPayloadController],
  providers: [PreviewPayloadService],
})
export class PreviewPayloadModule {}
