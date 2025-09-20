import { Controller, Get, Request, Query } from '@nestjs/common';
import { Cat9andcat12Service } from './cat9andcat12.service';
import { getUserId } from 'src/helper/common.helper';
import { Public } from 'src/decorators';

@Controller('cat9-and-cat12')
export class Cat9andcat12Controller {
  constructor(private readonly cat9andcat12Service: Cat9andcat12Service) {}

  // @Public()
  // @Get('get-data')
  // async getData(@Query('Date') date: string, @Request() req) {
  //   const userID = getUserId(req);
  //   const response = await this.cat9andcat12Service.getData(date, userID);

  //   return {
  //     statusCode: 200,
  //     message: 'Data List successfully!',
  //     data: response,
  //   };
  // }

  @Get('get-data')
  async getData(
    @Query('Date') date: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
  ) {
    return await this.cat9andcat12Service.getData(date, page, limit);
  }

  @Get('get-data-test')
  async getDataTest(
    @Query('date') date: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
  ) {
    return await this.cat9andcat12Service.getDataTest(date, page, limit);
  }

  @Get('test')
  async getTest(
    @Query('page') page: number,
    @Query('limit') limit: number,
  ) {
    return await this.cat9andcat12Service.getTest(page, limit);
  }
}
