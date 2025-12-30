import { Controller, Get, Post, Request, UseGuards } from '@nestjs/common';
import { LocalAuthGuard } from './guards/local-auth.guard';
import { AuthService } from './auth.service';
import { Public } from 'src/decorators';

@Controller('auth')
export class AuthController {
  constructor(private authService: AuthService) {}

  @Public()
  @UseGuards(LocalAuthGuard)
  @Post('login')
  async login(@Request() req) {
    const response = await this.authService.login(req.user);
    return response;
  }

  // @UseGuards(JwtAuthGuard)
  @Public()
  @Get('profile')
  getProfile(@Request() req) {
    return req.user;
  }
}
