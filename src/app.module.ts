import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { pgPoolProvider } from './pg.pool.provider';
@Module({
  imports: [],
  controllers: [AppController],
  providers: [AppService, pgPoolProvider],
})
export class AppModule {
}