import { Controller, Get, Query, Request } from '@nestjs/common';
import { FilemanagementService } from './filemanagement.service';
import { getUserId } from 'src/helper/common.helper';
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
}
