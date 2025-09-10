import { PartialType } from '@nestjs/mapped-types';
import { CreateCat9andcat12Dto } from './create-cat9andcat12.dto';

export class UpdateCat9andcat12Dto extends PartialType(CreateCat9andcat12Dto) {}
