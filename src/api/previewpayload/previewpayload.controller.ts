import { Controller, Get, Query, Request } from '@nestjs/common';
import { PreviewpayloadService } from './previewpayload.service';
import { getUserId } from 'src/helper/common.helper';

@Controller('previewpayload')
export class PreviewpayloadController {
  constructor(private readonly previewpayloadService: PreviewpayloadService) {}

  @Get('preview-excel')
  async previewExcel(
    @Query('module') module: string,
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('dockeyCMS') dockeyCMS: string,
    @Request() req,
  ) {
    const userId = getUserId(req);
    const res = await this.previewpayloadService.previewExcel(
      module,
      dateFrom,
      dateTo,
      factory,
      userId,
      dockeyCMS,
    );

    if (!res) return { statusCode: 401, message: 'Error export!' };
    return {
      statusCode: 200,
      message: 'Export success. Please wait for 5 minutes!',
    };
  }
}
