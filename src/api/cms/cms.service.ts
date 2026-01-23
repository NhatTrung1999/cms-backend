import { HttpService } from '@nestjs/axios';
import { HttpException, HttpStatus, Injectable } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { AxiosError } from 'axios';
import { catchError, firstValueFrom } from 'rxjs';

@Injectable()
export class CmsService {
  constructor(
    private readonly httpService: HttpService,
    private readonly configService: ConfigService,
  ) {}

  async createDataIntegrate(payload: any) {
    const externalApiUrl = this.configService.get('CMS_URL');

    const { data } = await firstValueFrom(
      this.httpService.post(externalApiUrl, payload).pipe(
        catchError((error: AxiosError) => {
          console.error(
            'Lỗi khi gọi BPM:',
            error.response?.data || error.message,
          );
          throw new HttpException(
            error.response?.data || 'Lỗi kết nối đến Server BPM',
            error.response?.status || HttpStatus.INTERNAL_SERVER_ERROR,
          );
        }),
      ),
    );
    return data;
    // return 'dfbdjsb'
  }
}
