import { HttpException, HttpStatus, Injectable } from '@nestjs/common';
import { CalculationResult, User } from '@prisma/client';
import { PrismaService } from 'src/prisma.service';
import { SignUpDto } from 'src/user/dto/signUp.dto';
import * as bcrypt from 'bcrypt';

@Injectable()
export class AdminService {
  constructor(private readonly prisma: PrismaService) {}

  async createUserByAdmin(dto: SignUpDto): Promise<{ message }> {
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
      data: { ...dto, password: hashedPassword, isConfirmed: true },
    });
    return { message: 'User created' };
  }

  async getAllUsers(): Promise<User[]> {
    const users = await this.prisma.user.findMany({});
    return users;
  }

  async updateUserinfo(id: number, dto: SignUpDto): Promise<{ message }> {
    const existingUser = await this.prisma.user.findUnique({
      where: {
        id: id,
      },
    });
    if (!existingUser) {
      throw new HttpException('user not found', HttpStatus.BAD_REQUEST);
    }
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
    const updatedUser = await this.prisma.user.update({
      where: { id: id },
      data: {
        ...dto,
        password: hashedPassword,
      },
    });
    return { message: 'user updated' };
  }

  async confirmedUser(id: number, dto: SignUpDto): Promise<{ message }> {
    const currentUser = await this.prisma.user.findUnique({
      where: {
        id: id,
      },
    });
    if (!currentUser) {
      throw new HttpException('user not found', HttpStatus.BAD_REQUEST);
    }
    const confirmedUser = await this.prisma.user.update({
      where: { id: id },
      data: { ...dto },
    });
    return { message: 'user confirmed' };
  }

  async getAllCalculationResult(): Promise<CalculationResult[]> {
    return await this.prisma.calculationResult.findMany({});
  }
}
