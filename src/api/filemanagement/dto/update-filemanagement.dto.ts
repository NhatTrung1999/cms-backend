import { PartialType } from '@nestjs/mapped-types';
import { CreateFilemanagementDto } from './create-filemanagement.dto';

export class UpdateFilemanagementDto extends PartialType(CreateFilemanagementDto) {}
