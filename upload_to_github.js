const fs = require('fs');
const https = require('https');
const path = require('path');

const token = process.env.GH_TOKEN;
if (!token) {
  console.error('ERROR: GH_TOKEN not set');
  process.exit(1);
}
console.log('Token length:', token.length);

const owner = 'Christinasaki';
const repo = 'ppt-translator';
const branch = 'main';
const baseDir = 'D:/06 AI编程实验/01 实验测试/翻译qoder';

const files = [
  '.gitignore',
  'README.md',
  'ppt_translate_standalone.py',
  'ppt一键翻译.html',
  'ppt-after-en.png',
  'ppt-before-zh.png'
];

function uploadFile(filePath, base64Content, message) {
  return new Promise((resolve) => {
    const encodedPath = encodeURIComponent(filePath);
    const data = JSON.stringify({
      message: message,
      content: base64Content,
      branch: branch
    });

    const options = {
      hostname: 'api.github.com',
      path: `/repos/${owner}/${repo}/contents/${encodedPath}`,
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/vnd.github.v3+json',
        'User-Agent': 'ppt-translator-upload',
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(data)
      }
    };

    const req = https.request(options, (res) => {
      let body = '';
      res.on('data', (chunk) => body += chunk);
      res.on('end', () => {
        const status = res.statusCode;
        if (status >= 200 && status < 300) {
          let sha = '';
          try {
            const json = JSON.parse(body);
            sha = json.commit ? json.commit.sha.substring(0, 7) : 'unknown';
          } catch (e) {}
          resolve(`OK: ${filePath} (${status}) sha=${sha}`);
        } else {
          resolve(`FAIL: ${filePath} (${status}): ${body.substring(0, 200)}`);
        }
      });
    });

    req.on('error', (e) => {
      resolve(`ERROR: ${filePath}: ${e.message}`);
    });

    req.write(data);
    req.end();
  });
}

async function main() {
  console.log('Starting upload...');
  for (const file of files) {
    const fullPath = path.join(baseDir, file);
    try {
      const content = fs.readFileSync(fullPath).toString('base64');
      const sizeKB = (content.length / 1024).toFixed(1);
      console.log(`Uploading ${file} (${sizeKB}KB)...`);
      const result = await uploadFile(file, content, `add ${file}`);
      console.log(result);
    } catch (e) {
      console.log(`SKIP: ${file} - ${e.message}`);
    }
  }
  console.log('=== Done ===');
}

main().catch(e => console.error('FATAL:', e));
