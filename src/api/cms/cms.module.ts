import { Module } from '@nestjs/common';
import { CmsService } from './cms.service';
import { CmsController } from './cms.controller';
import { HttpModule } from '@nestjs/axios';

@Module({
  imports: [
    HttpModule.register({
      timeout: 60 * 60 * 1000,
      maxRedirects: 5,
    }),
  ],
  controllers: [CmsController],
  providers: [CmsService],
})
export class CmsModule {}
