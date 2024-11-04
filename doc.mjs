'strict'
import Router from '@koa/router';
import { Canvas } from 'canvas';
import Table2canvas from 'table2canvas';
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import ImageModule from 'docxtemplater-image-module-free';
import fs from "fs";
import path from "path";
import sizeOf from "image-size";
import data from './data.js';

function router() {
  const api = new Router();

  api.get('/data', async ctx => {
    try {
      ctx.body = data;
    } catch (e) {
      console.error(e);
      ctx.body = e;
    }
  });

  api.post('/create', async ctx => {

    try {

      if (!ctx.request.body || !ctx.request.body.img1 || !ctx.request.body.img2 || !ctx.request.body.img3) throw new Error('缺失参数');

      const jingqing = ctx.request.body.jq; // text
      const image1 = ctx.request.body.img1; // base64
      const image2 = ctx.request.body.img2; // base64
      const image3 = ctx.request.body.img3; // json

      if (!Array.isArray(image3)) throw new Error('缺失参数');
      image3.shift();
      image3.shift();

      const table = new Table2canvas({
        canvas: new Canvas(0, 0),
        columns: columns,
        dataSource: image3,
        bgColor: '#fff',
        text: '各区网商投诉、建议类数据统计表',
        style: {
          fontSize: '24px',
          columnWidth: 120,
          borderColor: '#000000',
          padding: 10,
        }
      });
      const img1buff = Buffer.from(image1.split(',')[1], 'base64');
      const img2buff = Buffer.from(image2.split(',')[1], 'base64');
      const img3buff = table.canvas.toBuffer();

      const opts = {
        getImage: function (tagValue, tagName, meta) {
          if (tagName == 'image1') return img1buff;
          if (tagName == 'image2') return img2buff;
          if (tagName == 'image3') return img3buff;
          return fs.readFileSync(tagValue);
        },
        getSize: function (img, tagValue, tagName) {
          if (tagName == 'image3') return [670, 370];
          const dimensions = sizeOf(Buffer.from(img, "binary"));
          const height = 670 / dimensions.width * dimensions.height;
          return [670, height];
        }
      }
      const zip = new PizZip(fs.readFileSync(path.resolve("input.docx"), "binary"));

      const doc = new Docxtemplater(zip, {
        modules: [new ImageModule(opts)],
        linebreaks: true
      });

      console.log(jingqing);
      doc.render({
        jingqing: jingqing,
        image1: './demo.png',
        image2: './demo.png',
        image3: './demo.png'
      });

      const buf = doc.getZip().generate({
        type: "nodebuffer",
        compression: "DEFLATE",
      });
      fs.writeFileSync(path.resolve("./static/", "output.docx"), buf);

      ctx.body = { code: 'ok' };

    } catch (e) {
      console.error(e);
      ctx.body = e;
    }

  });

  return api;
}

const columns = [
  { title: '单位/项目', dataIndex: 'depart', textAlign: 'center' },
  {
    title: '投诉、建议类数据（去重复、去老用户）',
    textColor: 'rgba(255,0,0,1)',
    color: 'rgba(255,0,0,1)',
    titleFontWeight: 'bold',
    children: [
      { title: '本期', dataIndex: 'xszabq', textAlign: 'center' },
      { title: '环比', dataIndex: 'xszahb', textAlign: 'center' },
      { title: '同比', dataIndex: 'xszatb', textAlign: 'center' }
    ]
  },
  {
    title: '投诉数据（去重复、去老用户）',
    children: [
      { title: '本期', dataIndex: 'xsjqbq', textAlign: 'center' },
      { title: '环比', dataIndex: 'xsjqhb', textAlign: 'center' },
      { title: '同比', dataIndex: 'xsjqtb', textAlign: 'center' }
    ]
  },
  {
    title: '程序出错（含建议）网商数据（去重复、去老用户）',
    children: [
      { title: '本期', dataIndex: 'dqbq', textAlign: 'center' },
      { title: '环比', dataIndex: 'bqhb', textAlign: 'center' },
      { title: '同比', dataIndex: 'bqtb', textAlign: 'center' },
    ]
  },
  {
    title: '服务态度差（含建议）网商数据（去重复、去老用户）',
    children: [
      { title: '本期', dataIndex: 'zpbq', textAlign: 'center' },
      { title: '环比', dataIndex: 'zphb', textAlign: 'center' },
      { title: '同比', dataIndex: 'zptb', textAlign: 'center' },
    ]
  },
  {
    title: '质量问题（去重复、去老用户）',
    children: [
      { title: '本期', dataIndex: 'shbq', textAlign: 'center' },
      { title: '环比', dataIndex: 'shtb', textAlign: 'center' },
      { title: '同比', dataIndex: 'shhb', textAlign: 'center' },
    ]
  },
  {
    title: '沟通问题（去重复、去老用户）',
    children: [
      { title: '本期', dataIndex: 'odbq', textAlign: 'center' },
      { title: '环比', dataIndex: 'odhb', textAlign: 'center' },
      { title: '同比', dataIndex: 'odtb', textAlign: 'center' },
    ]
  },
  {
    title: '价格问题（去重复、去老用户）',
    children: [
      { title: '本期', dataIndex: 'xxbq', textAlign: 'center' },
      { title: '环比', dataIndex: 'xxhb', textAlign: 'center' },
      { title: '同比', dataIndex: 'xxtb', textAlign: 'center' },
    ]
  },
];

export default router;
