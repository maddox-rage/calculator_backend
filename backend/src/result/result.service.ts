import { HttpException, HttpStatus, Injectable } from '@nestjs/common';
import { CalculationResult } from '@prisma/client';
import { AuthService } from 'src/auth/auth.service';
import { PrismaService } from 'src/prisma.service';
import { writeResult } from './result.dto';
import { Response } from 'express';
import * as ExcelJS from 'exceljs';

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
        value3: dto.value3,
        value2: dto.value2,
        value1: dto.value1,
        uncertaintyBType: dto.uncertaintyBType,
        uncertaintyTotal: dto.uncertaintyTotal,
        uncertaintyExpanded: dto.uncertaintyExpanded,
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
  async getExcelCalculationResult(
    userId: number,
    calculatorId: number,
    res: Response,
  ): Promise<void> {
    const results = await this.prisma.calculationResult.findMany({
      where: {
        userId: userId,
        calculatorId: calculatorId,
      },
    });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Calculation Results');

    worksheet.columns = [
      { header: 'ID', key: 'id', width: 10 },
      { header: 'Результат', key: 'resultValue', width: 15 },
      { header: 'Разрядность', key: 'value1', width: 15 },
      { header: 'Результат измерений X', key: 'value2', width: 15 },
      { header: 'Абсолютная погрешность [Δ]', key: 'value3', width: 15 },
      {
        header: 'Неопределённость по типу В',
        key: 'uncertaintyBType',
        width: 20,
      },
      {
        header: 'Суммарная неопределённость',
        key: 'uncertaintyTotal',
        width: 20,
      },
      {
        header: 'Расширенная неопределённость',
        key: 'uncertaintyExpanded',
        width: 20,
      },
    ];

    results.forEach((result) => {
      worksheet.addRow(result);
    });

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename=' + 'calculation_results.xlsx',
    );

    await workbook.xlsx.write(res);
    res.end();
  }
}
