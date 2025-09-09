import { Injectable } from '@nestjs/common';
import { CreateFilemanagementDto } from './dto/create-filemanagement.dto';
import { UpdateFilemanagementDto } from './dto/update-filemanagement.dto';

@Injectable()
export class FilemanagementService {
  findAll() {
    return `This action returns all filemanagement`;
  }
}
