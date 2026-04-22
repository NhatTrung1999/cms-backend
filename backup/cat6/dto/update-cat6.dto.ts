import { PartialType } from '@nestjs/mapped-types';
import { CreateCat6Dto } from './create-cat6.dto';

export class UpdateCat6Dto extends PartialType(CreateCat6Dto) {}
