import { Body, Controller, Post } from '@nestjs/common';
import { CmsService } from './cms.service';

@Controller('cms')
export class CmsController {
  constructor(private readonly cmsService: CmsService) {}

  @Post('create')
  async create(@Body() body: any) {
    return await this.cmsService.createDataIntegrate(body);
  }
}
