export interface DecodedToken {
  email: string;
  login: string;
  sub: number;
  isAdmin: boolean;
  isConfirmed: boolean;
  iat: number;
  exp: number;
}
