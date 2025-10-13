import { Controller, Get } from '@nestjs/common';
import { Cat6Service } from './cat6.service';

@Controller('cat6')
export class Cat6Controller {
  constructor(private readonly cat6Service: Cat6Service) {}

  @Get('get-data-cat6')
  async getDataCat6() {
    return this.cat6Service.getDataCat6();
  }
}
