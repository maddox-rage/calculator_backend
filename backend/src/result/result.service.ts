import { HttpException, HttpStatus, Injectable } from '@nestjs/common';
import { CalculationResult } from '@prisma/client';
import { AuthService } from 'src/auth/auth.service';
import { PrismaService } from 'src/prisma.service';
import { writeResult } from './result.dto';

@Injectable()
export class ResultService {
  constructor(
    private readonly prisma: PrismaService,
    private readonly authService: AuthService,
  ) {}

  async getUserCalculationResults(
    userId: number,
  ): Promise<CalculationResult[]> {
    const existingUser = await this.prisma.user.findUnique({
      where: {
        id: userId,
      },
    });
    if (!existingUser) {
      throw new HttpException('User not found', HttpStatus.BAD_REQUEST);
    }
    return await this.prisma.calculationResult.findMany({
      where: {
        userId: userId,
      },
    });
  }

  async writeResult(userId: number, dto: writeResult): Promise<{ message }> {
    const existingUser = await this.prisma.user.findUnique({
      where: {
        id: userId,
      },
    });
    if (!existingUser) {
      throw new HttpException('User not found', HttpStatus.BAD_REQUEST);
    }
    await this.prisma.calculationResult.create({
      data: {
        resultValue: dto.resultValue,
        calculator: {
          connect: {
            id: dto.calculatorId,
          },
        },
        user: {
          connect: {
            id: userId,
          },
        },
      },
    });
    return { message: 'result writed' };
  }
  async getCalculationResultByUserIdAndCalculatorId(
    userId: number,
    calculatorId: number,
  ): Promise<CalculationResult[]> {
    const results = await this.prisma.calculationResult.findMany({
      where: {
        userId: userId,
        calculatorId: calculatorId,
      },
    });
    return results;
  }
}
