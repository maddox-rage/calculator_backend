import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { hostname } from 'os';

async function bootstrap() {
  const app = await NestFactory.create(AppModule);
  app.setGlobalPrefix('api');
  await app.listen(80, '0.0.0.0');
}
bootstrap();
