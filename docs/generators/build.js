/**
 * build.js
 * docs/ 配下のドキュメントをまとめて生成する。
 *   実行: cd docs/generators && npm install && npm run build
 */
const path = require('path');
const C = require('./common');

const OUT_DIR = path.join(__dirname, '..');

const DOCUMENTS = [
  { module: './spec', file: '仕様書_v2.3.docx' },
  { module: './usermanual', file: 'ユーザー用マニュアル_v2.3.docx' },
  { module: './setup', file: '初期設定マニュアル_v2.3.docx' },
  { module: './testchecklist', file: 'ユーザーテスト_チェックリスト_v2.3.docx' },
];

(async () => {
  for (const doc of DOCUMENTS) {
    const build = require(doc.module);
    await C.save(build(), path.join(OUT_DIR, doc.file));
  }
})();
