import { PartialType } from '@nestjs/mapped-types';
import { CreateCat1andcat4Dto } from './create-cat1andcat4.dto';

export class UpdateCat1andcat4Dto extends PartialType(CreateCat1andcat4Dto) {}
