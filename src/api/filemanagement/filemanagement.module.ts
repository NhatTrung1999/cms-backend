import { Module } from '@nestjs/common';
import { FilemanagementService } from './filemanagement.service';
import { FilemanagementController } from './filemanagement.controller';
import { DatabaseModule } from 'src/database/database.module';

@Module({
  imports: [DatabaseModule],
  controllers: [FilemanagementController],
  providers: [FilemanagementService],
})
export class FilemanagementModule {}
