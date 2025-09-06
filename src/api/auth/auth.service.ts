import { Injectable } from '@nestjs/common';
import { UsersService } from '../users/users.service';
import { JwtService } from '@nestjs/jwt';

@Injectable()
export class AuthService {
  constructor(
    private usersService: UsersService,
    private jwtService: JwtService,
  ) {}

  async validateUser(userid: string, pass: string, factory: string): Promise<any> {
    const user = await this.usersService.validateUser(userid, pass, factory);
    if (user && user.Password === pass) {
      const { Password, ...result } = user;
      return result;
    }
    return null;
  }

  async login(user: any) {
    const payload = { ...user };
    return {
      payload,
      access_token: this.jwtService.sign(payload),
    };
  }
}
