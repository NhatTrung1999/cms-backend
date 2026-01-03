import {
  BadRequestException,
  Controller,
  Get,
  Post,
  Query,
  Request,
  UploadedFile,
  UseInterceptors,
} from '@nestjs/common';
import { HrService } from './hr.service';
import { FileInterceptor } from '@nestjs/platform-express';
import { getFactortyID, getUserId } from 'src/helper/common.helper';

@Controller('hr')
export class HrController {
  constructor(private readonly hrService: HrService) {}

  @Get()
  async findAll(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.hrService.findAll(
      dateFrom,
      dateTo,
      factory,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
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
}
