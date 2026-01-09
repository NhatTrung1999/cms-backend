import { Module } from '@nestjs/common';
import { LogcatService } from './logcat.service';
import { LogcatController } from './logcat.controller';
import { DatabaseModule } from 'src/database/database.module';

@Module({
  imports: [DatabaseModule],
  controllers: [LogcatController],
  providers: [LogcatService],
})
export class LogcatModule {}
