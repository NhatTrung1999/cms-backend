import { Injectable } from '@nestjs/common';
import { UsersService } from '../users/users.service';
import { JwtService } from '@nestjs/jwt';

@Injectable()
export class AuthService {
  constructor(
    private usersService: UsersService,
    private jwtService: JwtService,
  ) {}

  async validateUser(
    userid: string,
    pass?: string,
    factory?: string,
  ): Promise<any> {
    // console.log(userid);
    const user = await this.usersService.validateUser(userid, pass, factory);
    // console.log(pass);
    if (pass){
      if (user && user.Password === pass) {
        const { Password, ...result } = user;
        return result;
      }
    } else {
      // if (user && user.Password === pass) {
        const { Password, ...result } = user;
        return result;
      // }
    }
    
    // return null;
  }

  async validateErpUser(userid: string, pass: string) {
    const user = await this.usersService.validateErpUser(userid, pass);
    if (user && user.PWD === pass) {
      const { PWD, ...result } = user;
      let res = await this.validateUser(result.USERID);
      return res
    }
    return null;
  }

  async checkLockErp(userid: string) {
    return await this.usersService.checkLock(userid);
  }

  async checkExistUser(userid: string){
    return await this.usersService.checkExistUser(userid)
  }

  async login(user: any) {
    const payload = { ...user };
    return {
      payload,
      access_token: this.jwtService.sign(payload),
    };
  }
}
