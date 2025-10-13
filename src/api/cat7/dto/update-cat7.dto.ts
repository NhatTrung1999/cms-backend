import { PartialType } from '@nestjs/mapped-types';
import { CreateCat7Dto } from './create-cat7.dto';

export class UpdateCat7Dto extends PartialType(CreateCat7Dto) {}
