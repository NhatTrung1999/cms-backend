import { PartialType } from '@nestjs/mapped-types';
import { CreateInfofactoryDto } from './create-infofactory.dto';

export class UpdateInfofactoryDto extends PartialType(CreateInfofactoryDto) {}
