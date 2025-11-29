import { Controller, Get, Query } from '@nestjs/common';
import { InfofactoryService } from './infofactory.service';

@Controller('infofactory')
export class InfofactoryController {
  constructor(private readonly infofactoryService: InfofactoryService) {}

  @Get('get-info-factory')
  async getInfoFactory(
    @Query('companyName') companyName: string,
    @Query('city') city: string,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.infofactoryService.getInfoFactory(
      companyName,
      city,
      sortField,
      sortOrder,
    );
  }
}

