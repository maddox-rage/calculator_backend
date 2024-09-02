import { Injectable, UnauthorizedException } from '@nestjs/common';
import { JwtService } from '@nestjs/jwt';
import { User } from '@prisma/client';
import { DecodedToken } from './token.interface';

@Injectable()
export class AuthService {
  constructor(private readonly jwtService: JwtService) {}

  async generateToken(user: User) {
    const payload = {
      email: user.email,
      sub: user.id,
      isAdmin: user.isAdmin,
      login: user.login,
    };
    return { token: this.jwtService.sign(payload) };
  }

  async verifyToken(token: string): Promise<DecodedToken> {
    try {
      return this.jwtService.verify<DecodedToken>(token);
    } catch (e) {
      throw new UnauthorizedException('Invalid or expired token');
    }
  }
}
