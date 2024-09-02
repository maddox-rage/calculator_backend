import {
  Body,
  Controller,
  Param,
  Patch,
  Post,
  UseGuards,
} from '@nestjs/common';
import { UserService } from './user.service';
import { SignUpDto } from './dto/signUp.dto';
import { AuthResponseDto } from 'src/auth/auth.dto';
import { SignInDto } from './dto/signIn.dto';
import { User } from '@prisma/client';
import { JwtAuthGuard } from 'src/guards/jwtAuth.guard';

@Controller('user')
export class UserController {
  constructor(private readonly userService: UserService) {}

  @Post('/signup')
  async signUp(@Body() userData: SignUpDto): Promise<AuthResponseDto> {
    return this.userService.signUp(userData);
  }

  @Post('/signin')
  async signIn(@Body() userData: SignInDto): Promise<AuthResponseDto> {
    return this.userService.signIn(userData);
  }
  @UseGuards(JwtAuthGuard)
  @Patch(':id')
  async updateUserInfo(
    @Param('id') id: number,
    @Body() userData: SignUpDto,
  ): Promise<User> {
    return this.userService.updateUserInfo(Number(id), userData);
  }
}
