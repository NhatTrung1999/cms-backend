import { Controller, Get } from '@nestjs/common';
import { Cat5Service } from './cat5.service';

@Controller('cat5')
export class Cat5Controller {
  constructor(private readonly cat5Service: Cat5Service) {}

  @Get()
  async getDataWMS() {
    return this.cat5Service.getDataWMS();
  }
}
