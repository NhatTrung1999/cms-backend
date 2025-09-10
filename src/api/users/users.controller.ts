import {
  Body,
  Controller,
  Delete,
  Get,
  Param,
  Patch,
  Post,
  Query,
} from '@nestjs/common';
import { UsersService } from './users.service';
import { CreateUserDto } from './dto/create-user.dto';
import { UpdateUserDto } from './dto/update-user.dto';

@Controller('users')
export class UsersController {
  constructor(private usersService: UsersService) {}

  @Get()
  async getAll() {
    const res = await this.usersService.getAll();
    return res;
  }

  @Get('get-one')
  async getOne(@Query('userid') userid: string, @Query('name') name: string) {
    return await this.usersService.getOne(userid, name);
  }

  @Post('add-user')
  async addUser(@Body() createUserDto: CreateUserDto) {
    const res = await this.usersService.addUser(createUserDto);
    return {
      statusCode: 200,
      message: 'Add successfully!',
      data: res[0],
    };
  }

  @Patch('update-user')
  async updateUser(@Body() updateUserDto: UpdateUserDto) {
    await this.usersService.updateUser(updateUserDto);
    return {
      statusCode: 200,
      message: 'Update successfully!',
    };
  }

  @Delete(':id')
  async deleteUser(@Param('id') id: string) {
    await this.usersService.deleteUser(id);
    return {
      statusCode: 200,
      message: 'Delete successfully!',
    };
  }
}
