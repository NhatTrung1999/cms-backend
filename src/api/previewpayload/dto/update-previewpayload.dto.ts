import { PartialType } from '@nestjs/mapped-types';
import { CreatePreviewpayloadDto } from './create-previewpayload.dto';

export class UpdatePreviewpayloadDto extends PartialType(CreatePreviewpayloadDto) {}
