import {
  CanActivate,
  ExecutionContext,
  Injectable,
  ForbiddenException,
} from '@nestjs/common';
import { AuthService } from 'src/auth/auth.service';

@Injectable()
export class AdminGuard implements CanActivate {
  constructor(private readonly authService: AuthService) {}

  async canActivate(context: ExecutionContext): Promise<boolean> {
    const request = context.switchToHttp().getRequest();
    const authHeader = request.headers.authorization;

    if (!authHeader) {
      throw new ForbiddenException('Token not provided');
    }

    const token = authHeader.split(' ')[1];
    const decodedToken = await this.authService.verifyToken(token);

    if (!decodedToken.isAdmin) {
      throw new ForbiddenException('User is not an administrator');
    }

    return true;
  }
}
