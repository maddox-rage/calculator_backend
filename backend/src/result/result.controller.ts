import {
  Body,
  Controller,
  Get,
  Param,
  Post,
  Res,
  UseGuards,
} from '@nestjs/common';
import { ResultService } from './result.service';
import { CalculationResult } from '@prisma/client';
import { JwtAuthGuard } from 'src/guards/jwtAuth.guard';
import { writeResult } from './result.dto';
import { Response } from 'express';
import { GetResultsDto } from './dto/getResultsById.dto';

@Controller('user/:userId/result/')
export class ResultController {
  constructor(private readonly resultService: ResultService) {}

  @UseGuards(JwtAuthGuard)
  @Get()
  async getUserCalculationResults(
    @Param('userId') userId: number,
  ): Promise<CalculationResult[]> {
    return this.resultService.getUserCalculationResults(Number(userId));
  }
  @UseGuards(JwtAuthGuard)
  @Post()
  async writeResult(
    @Param('userId') userId: number,
    @Body() resultData: writeResult,
  ): Promise<{ message }> {
    return this.resultService.writeResult(Number(userId), resultData);
  }
  @UseGuards(JwtAuthGuard)
  @Get(':calculatorId')
  async getCalculationResultByUserIdAndCalculatorId(
    @Param('userId') userId: number,
    @Param('calculatorId') calculatorId: number,
  ): Promise<CalculationResult[]> {
    return this.resultService.getCalculationResultByUserIdAndCalculatorId(
      Number(userId),
      Number(calculatorId),
    );
  }
  @Get('/:calculatorId/excel')
  async getExcelResult(
    @Param('userId') userId: number,
    @Param('calculatorId') calculatorId: number,
    @Res() res: Response,
  ): Promise<void> {
    await this.resultService.getExcelCalculationResult(
      Number(userId),
      Number(calculatorId),
      res,
    );
  }
  @Get('/:calculatorId/excel/byId')
  async getExcelResultByIDs(
    @Param('calculatorId') calculatorId: number,
    @Res() res: Response,
    @Body() resultIds: GetResultsDto,
  ): Promise<void> {
    await this.resultService.getExcelCalculationResultByResultId(
      resultIds,
      Number(calculatorId),
      res,
    );
  }
}
