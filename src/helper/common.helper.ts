import { isNullOrUndefined } from './util.helper';
import { Request } from 'express';
import { v5 as uuidv5 } from 'uuid';
import { createHash } from 'crypto';

export function getUserId(req: Request) {
  const user = getUser(req);
  return user?.userid ?? null;
}

export function getUser(req: Request) {
  if (isNullOrUndefined(req?.user)) {
    return null;
  }
  return <any>req?.user;
}

export function getDepartmentID(req: Request) {
  const user = getDeparment(req);
  return user?.department ?? null;
}

export function getDeparment(req: Request) {
  if (isNullOrUndefined(req?.user)) {
    return null;
  }
  return <any>req?.user;
}

export function getFactortyID(req: Request) {
  const user = getFactory(req);
  return user?.factory ?? null;
}

export function getFactory(req: Request) {
  if (isNullOrUndefined(req?.user)) {
    return null;
  }
  return <any>req?.user;
}

export function Convert(originalname: string) {
  const NAMESPACE = '123e4567-e89b-12d3-a456-426614174000'; // Thay thế bằng UUID của bạn nếu cần

  const hash = createHash('sha1').update(originalname).digest('hex');

  const bytes = Buffer.from(hash.slice(0, 32), 'hex');

  return uuidv5(bytes.toString('hex'), NAMESPACE);
}
