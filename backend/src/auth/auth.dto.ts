import { User } from '@prisma/client';

export class AuthResponseDto {
  user: User;
  token: string;
}
