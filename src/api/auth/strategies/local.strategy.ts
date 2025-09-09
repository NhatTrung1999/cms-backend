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
      passReqToCallback: true,
    });
  }

  async validate(req: Request): Promise<any> {
    const { userid, password, factory } = req.body;

    let user;

    if (userid === 'admin' || userid === 'ESG') {
      user = await this.authService.validateUser(userid, password, factory);
    } else {
      let checkLock = await this.authService.checkLockErp(userid);
      if (!checkLock) {
        throw new UnauthorizedException(
          'Account ERP is clocked! Please enter your other account!',
        );
      }
      const checkExist = await this.authService.checkExistUser(userid)

      if(!checkExist) {
        throw new UnauthorizedException(
          'This account has not been set up on the system!',
        );
      }

      user = await this.authService.validateErpUser(userid, password);
    }

    if (!user) {
      throw new UnauthorizedException('Account or password is not valid!');
    }
    return user;
  }
}
