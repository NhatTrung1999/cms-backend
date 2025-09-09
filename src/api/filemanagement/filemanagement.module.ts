import { Module } from '@nestjs/common';
import { FilemanagementService } from './filemanagement.service';
import { FilemanagementController } from './filemanagement.controller';

@Module({
  controllers: [FilemanagementController],
  providers: [FilemanagementService],
})
export class FilemanagementModule {}
