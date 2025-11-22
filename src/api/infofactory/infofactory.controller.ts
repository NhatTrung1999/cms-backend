import { Controller, Get } from '@nestjs/common';
import { InfofactoryService } from './infofactory.service';

@Controller('infofactory')
export class InfofactoryController {
  constructor(private readonly infofactoryService: InfofactoryService) {}

  @Get('get-info-factory')
  async getInfoFactory() {
    return this.infofactoryService.getInfoFactory();
  }
}
