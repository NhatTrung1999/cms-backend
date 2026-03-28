import { PartialType } from '@nestjs/mapped-types';
import { CreatePreviewPayloadDto } from './create-preview-payload.dto';

export class UpdatePreviewPayloadDto extends PartialType(CreatePreviewPayloadDto) {}
