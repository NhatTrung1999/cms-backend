import { Strategy } from 'passport-local';
import { PassportStrategy } from '@nestjs/passport';
import { Injectable, UnauthorizedException } from '@nestjs/common';
import { AuthService } from '../auth.service';
import { Request } from 'express';

@Injectable()
export class LocalStrategy extends PassportStrategy(Strategy) {
  constructor(private authService: AuthService) {
    super({
      usernameField: 'userid',
      passwordField: 'password',
      passReqToCallback: true
    });
  }

  async validate(req: Request): Promise<any> {
    // console.log(req.body);
    const { userid, password, factory } = req.body;
    const user = await this.authService.validateUser(userid, password, factory);
    if (!user) {
      throw new UnauthorizedException('Account or password is not valid!');
    }
    return user;
  }
}
