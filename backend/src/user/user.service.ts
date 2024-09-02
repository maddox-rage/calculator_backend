import { HttpException, HttpStatus, Injectable } from '@nestjs/common';
import { PrismaService } from 'src/prisma.service';
import { User } from '@prisma/client';
import { SignUpDto } from './dto/signUp.dto';
import { AuthService } from 'src/auth/auth.service';
import * as bcrypt from 'bcrypt';
import { SignInDto } from './dto/signIn.dto';

@Injectable()
export class UserService {
  constructor(
    private readonly prisma: PrismaService,
    private readonly authService: AuthService,
  ) {}

  async signUp(dto: SignUpDto): Promise<{ user: User; token: string }> {
    const existingUserWithThisLogin = await this.prisma.user.findUnique({
      where: {
        login: dto.login,
      },
    });
    const existingUserWithThisEmail = await this.prisma.user.findUnique({
      where: {
        email: dto.email,
      },
    });
    if (existingUserWithThisEmail) {
      throw new HttpException('email has be taken', HttpStatus.BAD_REQUEST);
    }

    if (existingUserWithThisLogin) {
      throw new HttpException('login has be taken', HttpStatus.BAD_REQUEST);
    }
    const hashedPassword = await bcrypt.hash(dto.password, 10);
    const user = await this.prisma.user.create({
      data: { ...dto, password: hashedPassword, isConfirmed: false },
    });
    const token = await this.authService.generateToken(user);
    return { user: user, ...token };
  }

  async signIn(dto: SignInDto): Promise<{ user: User; token: string }> {
    const curretUser = await this.prisma.user.findUnique({
      where: {
        login: dto.login,
      },
    });
    if (!curretUser) {
      throw new HttpException(
        'User with this login not found',
        HttpStatus.BAD_REQUEST,
      );
    }
    const isMatch = await bcrypt.compare(dto.password, curretUser.password);
    if (!isMatch) {
      throw new HttpException(
        'invalid login or password',
        HttpStatus.BAD_REQUEST,
      );
    }
    const token = await this.authService.generateToken(curretUser);
    return { user: curretUser, ...token };
  }

  async updateUserInfo(id: number, dto: SignUpDto): Promise<User> {
    const existingUserWithThisEmail = await this.prisma.user.findUnique({
      where: {
        email: dto.email,
      },
    });
    const existingUserWithThisLogin = await this.prisma.user.findUnique({
      where: {
        login: dto.login,
      },
    });
    if (existingUserWithThisEmail) {
      throw new HttpException('email has be taken', HttpStatus.BAD_REQUEST);
    }
    if (existingUserWithThisLogin) {
      throw new HttpException('login has be taken', HttpStatus.BAD_REQUEST);
    }
    const hashedPassword = await bcrypt.hash(dto.password, 10);
    const user = await this.prisma.user.update({
      where: { id: id },
      data: { ...dto, password: hashedPassword },
    });
    return user;
  }
}
