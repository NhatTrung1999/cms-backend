import { Controller, Get, Query } from '@nestjs/common';
import { Cat1andcat4Service } from './cat1andcat4.service';

@Controller('cat1andcat4')
export class Cat1andcat4Controller {
  constructor(private readonly cat1andcat4Service: Cat1andcat4Service) {}

  @Get('get-data-cat1-and-cat4')
  async getDataWMS(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat1andcat4Service.getDataCat1AndCat4(
      dateFrom,
      dateTo,
      factory,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }
}
