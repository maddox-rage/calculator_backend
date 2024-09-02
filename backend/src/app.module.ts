import { Module } from '@nestjs/common';
import { UserModule } from './user/user.module';
import { AuthModule } from './auth/auth.module';
import { AdminModule } from './admin/admin.module';
import { ResultModule } from './result/result.module';

@Module({
  imports: [UserModule, AuthModule, AdminModule, ResultModule],
})
export class AppModule {}
