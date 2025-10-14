import { Controller, Get, Query } from '@nestjs/common';
import { Cat6Service } from './cat6.service';

@Controller('cat6')
export class Cat6Controller {
  constructor(private readonly cat6Service: Cat6Service) {}

  @Get('get-data-cat6')
  async getDataCat6(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat6Service.getDataCat6(
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
