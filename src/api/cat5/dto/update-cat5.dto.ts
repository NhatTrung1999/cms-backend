import { PartialType } from '@nestjs/mapped-types';
import { CreateCat5Dto } from './create-cat5.dto';

export class UpdateCat5Dto extends PartialType(CreateCat5Dto) {}
