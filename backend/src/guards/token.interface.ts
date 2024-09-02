export interface DecodedToken {
  email: string;
  sub: number;
  isAdmin: boolean;
  iat: number;
  exp: number;
}
