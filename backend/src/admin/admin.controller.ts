import {
  Body,
  Controller,
  Get,
  Param,
  Patch,
  Post,
  UseGuards,
} from '@nestjs/common';
import { AdminService } from './admin.service';
import { SignUpDto } from 'src/user/dto/signUp.dto';
import { CalculationResult, User } from '@prisma/client';
import { JwtAuthGuard } from 'src/guards/jwtAuth.guard';
import { AdminGuard } from 'src/guards/admin.guard';

@Controller('admin')
export class AdminController {
  constructor(private readonly adminService: AdminService) {}

  @UseGuards(JwtAuthGuard, AdminGuard)
  @Post('/createUser')
  async createUserByAdmin(
    @Body() userData: SignUpDto,
  ): Promise<{ message: string }> {
    return this.adminService.createUserByAdmin(userData);
  }

  @UseGuards(JwtAuthGuard, AdminGuard)
  @Get()
  async getAllUsers(): Promise<User[]> {
    return this.adminService.getAllUsers();
  }

  @UseGuards(JwtAuthGuard, AdminGuard)
  @Patch(':id')
  async updateUserInfo(
    @Param('id') id: number,
    @Body() userData: SignUpDto,
  ): Promise<{ message }> {
    return this.adminService.updateUserinfo(Number(id), userData);
  }

  @UseGuards(JwtAuthGuard, AdminGuard)
  @Patch('/confirmed/:id')
  async confirmedUser(
    @Param('id') id: number,
    @Body() userData: SignUpDto,
  ): Promise<{ message }> {
    return this.adminService.confirmedUser(Number(id), userData);
  }
  @UseGuards(JwtAuthGuard, AdminGuard)
  @Get('/results')
  async getAllCalcultionResults(): Promise<CalculationResult[]> {
    return this.adminService.getAllCalculationResult();
  }
}
