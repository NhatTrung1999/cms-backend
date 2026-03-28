import { Injectable } from '@nestjs/common';

@Injectable()
export class PreviewPayloadService {
  async previewPayloadType(
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockey: string,
  ) {
    return { dateFrom, dateTo, factory, dockey };
  }
}
