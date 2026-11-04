import { Request } from 'express';

/** The request of a controller method (`@Request() req`); JwtStrategy.validate() fills `req.user`. */
export type UserRequest = Request;

/** Fields that JwtStrategy.validate() puts on `req.user`. */
type RequestUser = {
  userid?: any;
  username?: any;
  factory?: any;
};

export function getUser(req: Request) {
  return (req?.user as RequestUser | undefined) ?? null;
}

export function getUserId(req: Request) {
  return getUser(req)?.userid ?? null;
}

export function getFactoryID(req: Request) {
  return getUser(req)?.factory ?? null;
}
