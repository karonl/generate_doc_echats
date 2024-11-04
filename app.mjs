'strict'
import http from 'http';
import doc from './doc.mjs';
import Koa from 'koa';
import Router from '@koa/router';
import serve from "koa-static";
import cors from 'koa2-cors';
import bodyParser from 'koa-bodyparser';

!async function () {
  const app = new Koa();
  const router = new Router();

  app.use(cors({ origin: '*' }));
  app.use(serve("./static"));

  router.use('/doc', doc().routes());
  app.use(bodyParser({
    jsonLimit: '100mb',
    formLimit: '100mb',
    textLimit: '100mb'
  }));
  app.use(router.routes());
  app.use(router.allowedMethods());

  http.createServer(app.callback()).listen(5000);
  console.log('===============');
  console.log('APaas started! http://::5000');
  console.log('===============');
}();
