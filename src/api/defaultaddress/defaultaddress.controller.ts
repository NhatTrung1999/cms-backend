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
import { DefaultaddressService } from './defaultaddress.service';

@Controller('defaultaddress')
export class DefaultaddressController {
  constructor(private readonly defaultaddressService: DefaultaddressService) {}

  @Get('get-default-address')
  async getDefaultAddress(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.defaultaddressService.getDefaultAddress(
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
