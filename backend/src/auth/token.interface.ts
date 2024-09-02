export interface DecodedToken {
  email: string;
  login: string;
  sub: number;
  isAdmin: boolean;
  iat: number;
  exp: number;
}
