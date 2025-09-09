import {
  Controller,
  Get,
  Post,
  Body,
  Patch,
  Param,
  Delete,
} from '@nestjs/common';
import { FilemanagementService } from './filemanagement.service';
import { CreateFilemanagementDto } from './dto/create-filemanagement.dto';
import { UpdateFilemanagementDto } from './dto/update-filemanagement.dto';

@Controller('filemanagement')
export class FilemanagementController {
  constructor(private readonly filemanagementService: FilemanagementService) {}

  @Get('get-data')
  findAll() {
    return this.filemanagementService.findAll();
  }
}
