import {
  Controller,
  HttpException,
  HttpStatus,
  Inject,
  Post,
  UploadedFile,
  UseInterceptors,
} from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express/multer';
import { constants } from './include';
import { Pool } from 'pg';
import { AppService } from './app.service';
import { ApiConsumes, ApiBody } from '@nestjs/swagger';

@Controller('excel')
export class AppController {
  constructor(
    private readonly appServices: AppService,
    @Inject('PG_POOL') private readonly pool: Pool,
  ) { }

  @Post('upload')
  @ApiConsumes('multipart/form-data')
  @ApiBody({
    schema: {
      type: 'object',
      properties: {
        file: {
          type: 'string',
          format: 'binary',
        },
      },
    },
  })
  @UseInterceptors(FileInterceptor('file'))

  async uploadFile(@UploadedFile() file: Express.Multer.File) {
    try {
      const result = await this.appServices.uploadFile(file);
      return result;
    } catch (error) {
      throw new HttpException(
        constants.serverError,
        HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

}