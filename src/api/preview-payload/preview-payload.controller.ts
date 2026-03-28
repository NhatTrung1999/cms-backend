import { Controller, Get, Query } from '@nestjs/common';
import { PreviewPayloadService } from './preview-payload.service';

@Controller('preview-payload')
export class PreviewPayloadController {
  constructor(private readonly previewPayloadService: PreviewPayloadService) {}

  @Get('preview-payload-type')
  async previewPayloadType(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('dockey') dockey: string,
  ) {
    return this.previewPayloadService.previewPayloadType(
      dateFrom,
      dateTo,
      factory,
      dockey,
    );
  }
}
