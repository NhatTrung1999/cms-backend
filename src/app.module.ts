import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { DatabaseModule } from './database/database.module';
import { ConfigModule } from '@nestjs/config';
import { AuthModule } from './api/auth/auth.module';
import { APP_GUARD } from '@nestjs/core';
import { JwtAuthGuard } from './api/auth/guards/jwt-auth.guard';
import { UsersModule } from './api/users/users.module';
import { FilemanagementModule } from './api/filemanagement/filemanagement.module';
import { Cat9andcat12Module } from './api/cat9andcat12/cat9andcat12.module';
import { EventsModule } from './api/events/events.module';
import { Cat5Module } from './api/cat5/cat5.module';
import { ServeStaticModule } from '@nestjs/serve-static';
import { join } from 'path';
import { Cat7Module } from './api/cat7/cat7.module';
import { Cat6Module } from './api/cat6/cat6.module';

@Module({
  imports: [
    ServeStaticModule.forRoot({
      rootPath: join(__dirname, '..', 'frontend', 'dist'),
    }),
    ConfigModule.forRoot({ isGlobal: true }),
    DatabaseModule,
    AuthModule,
    UsersModule,
    EventsModule,
    FilemanagementModule,
    Cat9andcat12Module,
    Cat5Module,
    Cat7Module,
    Cat6Module,
  ],
  controllers: [AppController],
  providers: [
    AppService,
    {
      provide: APP_GUARD,
      useClass: JwtAuthGuard,
    },
  ],
})
export class AppModule {}
