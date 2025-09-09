import { Body, Controller, Delete, Get, Param, Patch, Post, Query } from '@nestjs/common';
import { UsersService } from './users.service';
import { CreateUserDto } from './dto/create-user-dto';

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
    const res = await this.usersService.addUser(createUserDto)
    return {
      statusCode: 201,
      message: 'Add successfully!',
      data: res[0]
    };
  }

  @Patch(':id')
  async updateUser(@Param('id') id: string, @Body() updateUserDto: CreateUserDto){
    console.log(id, updateUserDto);
  }

  @Delete(':id')
  async deleteUser(@Param('id') id: string) {
    console.log(id);
  }
}
