export class CreateUserDto {
  id: string;
  userid: string;
  name: string;
  email?: string;
  role: string;
  status: boolean;
  createdAt?: string;
  updatedAt?: string;
}
