// Apps Script にアップロードせずに、ブラウザだけで画面を確認するためのファイルを作ります
// 使い方：このフォルダで  node tools/preview.mjs  → dev/preview.html をブラウザで開く
import { readFileSync, writeFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const src = join(root, "src");
const read = p => readFileSync(p, "utf8");

// <?!= include('js/store'); ?> を、src/js/store.html の中身に置き換える
let html = read(join(src, "index.html"))
  .replace(/<\?!=\s*include\('([^']+)'\);?\s*\?>/g, (_, name) => read(join(src, name + ".html")));

// サーバー側の .js をまとめ、公開関数（名前が _ で終わらない function）を window.__gas に登録する
const serverFiles = readdirSync(src).filter(f => f.endsWith(".js"));
const server = serverFiles.map(f => read(join(src, f))).join("\n");
const publicFns = [...server.matchAll(/^function\s+([A-Za-z0-9]+)\s*\(/gm)].map(m => m[1]).filter(n => !n.endsWith("_"));
const serverBlock = `(function(){\n${server}\nwindow.__gas = { ${publicFns.join(", ")} };\n})();`;

// <?= shopName ?> を、Code.js の SHOP_NAME の値に置き換える（本番では doGet が入れる）
const shopName = (server.match(/const SHOP_NAME = '([^']*)'/) || [, "SAMPLE SALON"])[1];
html = html.replace(/<\?=\s*shopName\s*\?>/g, shopName);

// viewport は本番では doGet で付けるので、プレビューではここで足す
html = html.replace("<head>", `<head>\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">`);
html = html.replace("<body>", `<body>\n<script>\n${read(join(root, "dev", "mock-gas.js")).replace(/<\/script/gi, "<\\/script")}\n</script>\n<script>\n${serverBlock.replace(/<\/script/gi, "<\\/script")}\n</script>`);

writeFileSync(join(root, "dev", "preview.html"), html);
console.log("dev/preview.html を作りました（公開関数：" + publicFns.join(", ") + "）");
