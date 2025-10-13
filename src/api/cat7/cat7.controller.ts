import {
  Controller,
  Get,
  Post,
  Body,
  Patch,
  Param,
  Delete,
  Query,
} from '@nestjs/common';
import { Cat7Service } from './cat7.service';
import { CreateCat7Dto } from './dto/create-cat7.dto';
import { UpdateCat7Dto } from './dto/update-cat7.dto';

@Controller('cat7')
export class Cat7Controller {
  constructor(private readonly cat7Service: Cat7Service) {}

  @Get('get-data-cat7')
  async getDataCat7(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat7Service.getDataCat7(
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
