import {
  Controller,
  Get,
  Post,
  Query,
  UseInterceptors,
  UploadedFile,
  Request,
  BadRequestException,
  Patch,
  Param,
  Body,
  Delete,
} from '@nestjs/common';
import { DefaultaddressService } from './defaultaddress.service';
import { FileInterceptor } from '@nestjs/platform-express';
import { getFactortyID, getUserId } from 'src/helper/common.helper';

@Controller('defaultaddress')
export class DefaultaddressController {
  constructor(private readonly defaultaddressService: DefaultaddressService) {}

  @Get('get-default-address')
  async getDefaultAddress(
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.defaultaddressService.getDefaultAddress(sortField, sortOrder);
  }

  @Patch('update-default-address/:id')
  async updateDefaultAddress(
    @Param('id') id: string,
    @Body()
    updateDto: {
      DefaultAddress: string;
    },
    @Request() req,
  ) {
    const factory = getFactortyID(req);
    const userid = getUserId(req);
    return this.defaultaddressService.updateDefaultAddress(
      id,
      updateDto,
      factory,
      userid,
    );
  }

  @Delete('delete-default-address/:id')
  async deleteDefaultAddress(@Param('id') id: string) {
    return this.defaultaddressService.deleteDefaultAddress(id);
  }

  @Post('import-excel-default-address')
  @UseInterceptors(
    FileInterceptor('file', { limits: { fileSize: 50 * 1024 * 1024 } }),
  )
  async importExcelDefaultAddress(
    @UploadedFile() file: Express.Multer.File,
    @Request() req,
  ) {
    const factory = getFactortyID(req);
    const userid = getUserId(req);
    console.log(file);
    if (!file) throw new BadRequestException('No file uploaded!');

    const data = await this.defaultaddressService.importExcelDefaultAddress(
      file,
      userid,
      factory,
    );
    return data;
  }

  @Post('sync-default-address')
  async syncDefaultAddress(
    @Body() body: { factory: string; syncDefaultAddress: string },
    @Request() req,
  ) {
    const userid = getUserId(req);
    return this.defaultaddressService.syncDefaultAddress(
      userid,
      body.factory,
      body.syncDefaultAddress,
    );
  }
}
