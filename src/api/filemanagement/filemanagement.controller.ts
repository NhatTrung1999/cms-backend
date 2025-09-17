import {
  Controller,
  Get,
  NotFoundException,
  Param,
  Query,
  Request,
  Res,
} from '@nestjs/common';
import { FilemanagementService } from './filemanagement.service';
import { getUserId } from 'src/helper/common.helper';
import { Response } from 'express';
import * as fs from 'fs';
import * as path from 'path';
// import { Public } from 'src/decorators';

@Controller('filemanagement')
export class FilemanagementController {
  constructor(private readonly filemanagementService: FilemanagementService) {}

  // @Public()
  @Get('get-data')
  async getData(
    @Query('Module') module: string,
    @Query('File_Name') file_name: string,
    @Request() req,
  ) {
    // console.log(module, file_name, req.user);
    const userID = getUserId(req);
    const response = await this.filemanagementService.getData(
      module,
      file_name,
      userID,
    );

    return {
      statusCode: 200,
      message: 'Data List successfully!',
      data: response,
    };
  }

  @Get('generate-file-excel')
  async generateFileExcel(
    @Query('Module') module: string,
    @Query('Date') date: string,
    @Request() req,
  ) {
    // console.log(module, date, req.user);
    const userID = getUserId(req);
    return await this.filemanagementService.generateFileExcel(
      module,
      date,
      userID,
    );
  }

  @Get('download/:id')
  async downloadFile(@Param('id') id: string, @Res() res: Response) {
    const fileRecord = await this.filemanagementService.getFileById(id);
    console.log(fileRecord);
    if (!fs.existsSync(fileRecord.Path)) {
      throw new NotFoundException('File not found on server');
    }
    return res.download(fileRecord.Path, path.basename(fileRecord.File_Name));
  }
}
