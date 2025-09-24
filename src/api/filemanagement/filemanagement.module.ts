import { Module } from '@nestjs/common';
import { FilemanagementService } from './filemanagement.service';
import { FilemanagementController } from './filemanagement.controller';
import { DatabaseModule } from 'src/database/database.module';
import { EventsModule } from '../events/events.module';

@Module({
  imports: [DatabaseModule, EventsModule],
  controllers: [FilemanagementController],
  providers: [FilemanagementService],
})
export class FilemanagementModule {}
