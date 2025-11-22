import { Module } from '@nestjs/common';
import { InfofactoryService } from './infofactory.service';
import { InfofactoryController } from './infofactory.controller';
import { DatabaseModule } from 'src/database/database.module';

@Module({
  imports: [DatabaseModule],
  controllers: [InfofactoryController],
  providers: [InfofactoryService],
})
export class InfofactoryModule {}
