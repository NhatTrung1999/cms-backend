import {
  BadRequestException,
  Body,
  Controller,
  Delete,
  Get,
  Param,
  Patch,
  Post,
  Query,
  Request,
  Res,
  StreamableFile,
  UploadedFile,
  UseInterceptors,
} from '@nestjs/common';
import { Cat1andcat4Service } from './cat1andcat4.service';
import { FileInterceptor } from '@nestjs/platform-express';
import { getFactortyID, getUserId } from 'src/helper/common.helper';
import { Public } from 'src/decorators';
import { Response } from 'express';
import dayjs from 'dayjs';
dayjs().format();

@Controller('cat1andcat4')
export class Cat1andcat4Controller {
  constructor(private readonly cat1andcat4Service: Cat1andcat4Service) {}

  @Get('get-data-cat1-and-cat4')
  async getDataCat1AndCat4(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('usage') usage: boolean,
    @Query('unitWeight') unitWeight: boolean,
    @Query('weight') weight: boolean,
    @Query('departure') departure: boolean,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat1andcat4Service.getDataCat1AndCat4(
      dateFrom,
      dateTo,
      factory,
      String(usage) === 'true',
      String(unitWeight) === 'true',
      String(weight) === 'true',
      String(departure) === 'true',
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }

  @Get('get-port-code')
  async getPortCode(
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat1andcat4Service.getPortCode(sortField, sortOrder);
  }

  @Get('get-tax-free-zone-address')
  async getTaxFreeZoneAddress(
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat1andcat4Service.getTaxFreeZoneAddress(sortField, sortOrder);
  }

  @Get('get-style-auto-fill')
  async getStyleAutoFill(
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat1andcat4Service.getStyleAutoFill(sortField, sortOrder);
  }

  @Delete('style-auto-fill/:id')
  async deleteStyleAutoFill(@Param('id') id: string) {
    return this.cat1andcat4Service.deleteStyleAutoFill(id);
  }

  @Patch('tax-free-zone-address/:id')
  async update(
    @Param('id') id: string,
    @Body()
    updateDto: {
      TaxFreeZoneAddress: string;
    },
    @Request() req,
  ) {
    const factory = getFactortyID(req);
    const userid = getUserId(req);
    return this.cat1andcat4Service.update(id, updateDto, factory, userid);
  }

  @Post('import-excel-port-code')
  @UseInterceptors(
    FileInterceptor('file', { limits: { fileSize: 50 * 1024 * 1024 } }),
  )
  async importExcelPortCode(
    @UploadedFile() file: Express.Multer.File,
    @Request() req,
  ) {
    const factory = getFactortyID(req);
    const userid = getUserId(req);
    if (!file) throw new BadRequestException('No file uploaded!');

    const data = await this.cat1andcat4Service.importExcelPortCode(
      file,
      userid,
      factory,
    );
    return data;
  }

  @Post('import-excel-style-auto-fill')
  @UseInterceptors(
    FileInterceptor('file', { limits: { fileSize: 50 * 1024 * 1024 } }),
  )
  async importExcelStyleAutoFill(
    @UploadedFile() file: Express.Multer.File,
    @Request() req,
  ) {
    const factory = getFactortyID(req);
    const userid = getUserId(req);
    if (!file) throw new BadRequestException('No file uploaded!');

    const data = await this.cat1andcat4Service.importExcelStyleAutoFill(
      file,
      userid,
      factory,
    );
    return data;
  }

  @Post('import-excel-tax-free-zone-address')
  @UseInterceptors(
    FileInterceptor('file', { limits: { fileSize: 50 * 1024 * 1024 } }),
  )
  async importExcelTaxFreeZoneAddress(
    @UploadedFile() file: Express.Multer.File,
    @Request() req,
  ) {
    const factory = getFactortyID(req);
    const userid = getUserId(req);
    if (!file) throw new BadRequestException('No file uploaded!');

    const data = await this.cat1andcat4Service.importExcelTaxFreeZoneAddress(
      file,
      userid,
      factory,
    );
    return data;
  }

  @Public()
  @Get('auto-sent-cms')
  async autoSentCMS(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('dockey') dockeyCMS: string,
  ) {
    const data = await this.cat1andcat4Service.autoSentCMS(
      dateFrom,
      dateTo,
      factory,
      dockeyCMS,
    );
    return data;
  }

  @Public()
  @Get('verification-report')
  async getVerificationReport(
    @Query('previewDateFrom') previewDateFrom: string,
    @Query('previewDateTo') previewDateTo: string,
    @Query('loggingDateFrom') loggingDateFrom: string,
    @Query('loggingDateTo') loggingDateTo: string,
    @Query('factory') factory: string,
    @Query('category') category: string,
    @Query('status') status: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
  ) {
    return this.cat1andcat4Service.getVerificationReport({
      previewDateFrom,
      previewDateTo,
      loggingDateFrom,
      loggingDateTo,
      factory,
      category,
      status,
      page: +page || 1,
      limit: +limit || 50,
    });
  }

  @Public()
  @Get('export-preview-payload')
  async exportPreviewPayload(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('dockey') dockeyCMS: string,
    @Res({ passthrough: true }) res: Response,
  ): Promise<StreamableFile> {
    const buffer = await this.cat1andcat4Service.exportPreviewPayload(
      dateFrom,
      dateTo,
      factory,
      dockeyCMS,
    );

    res.set({
      'Content-Type':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="PreviewPayload_${dayjs().format('YYYYMMDD_HHmmss')}.xlsx"`,
      'Content-Length': buffer.byteLength,
    });
    return new StreamableFile(Buffer.from(buffer));
  }
}
