import {
  BadRequestException,
  Body,
  Controller,
  Get,
  Param,
  Patch,
  Post,
  Query,
  Request,
  Res,
  UploadedFile,
  UseInterceptors,
} from '@nestjs/common';
import { HrService } from './hr.service';
import { FileInterceptor } from '@nestjs/platform-express';
import { getFactortyID, getUserId } from 'src/helper/common.helper';
import { Response } from 'express';

@Controller('hr')
export class HrController {
  constructor(private readonly hrService: HrService) {}

  @Get()
  async findAll(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('fullName') fullName: string,
    @Query('id') id: string,
    @Query('department') department: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
    @Request() req,
  ) {
    const factory = getFactortyID(req);
    return this.hrService.findAll(
      dateFrom,
      dateTo,
      fullName,
      id,
      department,
      factory,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }

  @Patch(':id')
  async update(
    @Param('id') id: string,
    @Body()
    updateDto: {
      CurrentAddress: string;
      TransportationMethod: string;
    },
    @Request() req,
  ) {
    const factory = getFactortyID(req);
    const userid = getUserId(req);
    return this.hrService.update(id, updateDto, factory, userid);
  }

  @Post('import')
  @UseInterceptors(FileInterceptor('file'))
  async importFromExcel(
    @UploadedFile() file: Express.Multer.File,
    @Request() req,
  ) {
    const factory = getFactortyID(req);
    const userid = getUserId(req);
    if (!file) throw new BadRequestException('No file uploaded!');
    return this.hrService.importFromExcel(file, userid, factory);
  }

  @Get('export')
  async exportToExcel(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('fullName') fullName: string,
    @Query('id') id: string,
    @Query('department') department: string,
    @Request() req,
    @Res() res: Response,
  ) {
    const factory = getFactortyID(req);
    const buffer = await this.hrService.exportToExcel(
      dateFrom,
      dateTo,
      fullName,
      id,
      department,
      factory,
    );
    res.header('Content-Disposition', 'attachment; filename="danh_sach.xlsx"');
    res.type(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );

    res.send(buffer);
  }
}
