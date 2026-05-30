import { Injectable, UnauthorizedException } from '@nestjs/common';
import { UsersService } from '../users/users.service';
import { JwtService } from '@nestjs/jwt';
import { ConfigService } from '@nestjs/config';

@Injectable()
export class AuthService {
  constructor(
    private usersService: UsersService,
    private jwtService: JwtService,
    private configService: ConfigService,
  ) {}

  async validateUser(
    userid: string,
    pass?: string,
    factory?: string,
  ): Promise<any> {
    const user = await this.usersService.validateUser(userid, pass, factory);
    if (!user) return null;
    const { Password, ...result } = user;
    return result;
  }

  async validateErpUser(userid: string, pass: string) {
    const user = await this.usersService.validateErpUser(userid, pass);
    if (!user) return null;
    const { PWD, ...result } = user;
    return this.validateUser(result.USERID);
  }

  async checkLockErp(userid: string) {
    return await this.usersService.checkLock(userid);
  }

  async checkExistUser(userid: string){
    return await this.usersService.checkExistUser(userid)
  }

  async login(user: any) {
    const permissionState = await this.usersService.getUserModulePermissionState(
      user.UserID,
    );
    const payload = { ...user, ...permissionState };
    return {
      payload,
      access_token: this.signAccessToken(payload),
      refresh_token: this.signRefreshToken(payload),
    };
  }

  async refresh(refreshToken: string) {
    if (!refreshToken) {
      throw new UnauthorizedException('Refresh token is required');
    }

    try {
      const decoded = await this.jwtService.verifyAsync(refreshToken, {
        secret: this.getRefreshSecret(),
      });

      if (decoded?.type !== 'refresh' || !decoded?.payload) {
        throw new UnauthorizedException('Invalid refresh token');
      }

      const payload = decoded.payload;
      return {
        payload,
        access_token: this.signAccessToken(payload),
        refresh_token: this.signRefreshToken(payload),
      };
    } catch (error) {
      if (error instanceof UnauthorizedException) throw error;
      throw new UnauthorizedException('Invalid or expired refresh token');
    }
  }

  private signAccessToken(payload: any) {
    return this.jwtService.sign(payload, {
      expiresIn: '15m',
    });
  }

  private signRefreshToken(payload: any) {
    return this.jwtService.sign(
      {
        type: 'refresh',
        payload,
      },
      {
        secret: this.getRefreshSecret(),
        expiresIn: '30d',
      },
    );
  }

  private getRefreshSecret() {
    return (
      this.configService.get<string>('JWT_REFRESH_SECRET') ||
      `${this.configService.get<string>('JWT_SECRET', 'defaultSecret')}_refresh`
    );
  }
}
