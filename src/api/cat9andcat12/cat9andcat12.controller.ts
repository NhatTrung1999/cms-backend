import { Controller, Get, Request, Query } from '@nestjs/common';
import { Cat9andcat12Service } from './cat9andcat12.service';
import { getUserId } from 'src/helper/common.helper';
import { Public } from 'src/decorators';

@Controller('cat9-and-cat12')
export class Cat9andcat12Controller {
  constructor(private readonly cat9andcat12Service: Cat9andcat12Service) {}

  @Get('get-data-cat9-and-cat12')
  async getTest(
    @Query('date') date: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return await this.cat9andcat12Service.getData(
      date,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }
}
