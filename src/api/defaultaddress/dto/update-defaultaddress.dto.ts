import { PartialType } from '@nestjs/mapped-types';
import { CreateDefaultaddressDto } from './create-defaultaddress.dto';

export class UpdateDefaultaddressDto extends PartialType(CreateDefaultaddressDto) {}
