import { Controller, Get, Query } from '@nestjs/common';
import { Cat5Service } from './cat5.service';

@Controller('cat5')
export class Cat5Controller {
  constructor(private readonly cat5Service: Cat5Service) {}

  @Get('get-data-cat5')
  async getDataWMS(
    @Query('date') date: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat5Service.getDataWMS(
      date,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }
}
