import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { exec } from 'child_process'

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    react(),
    {
      name: 'rag-reindex-api',
      configureServer(server) {
        server.middlewares.use(async (req, res, next) => {
          if (req.url === '/api/reindex' && req.method === 'POST') {
            const pythonCmd = 'python';
            const fs = await import('fs');
            const path = await import('path');
            
            // 다중 기기(집/사무실 노트북) 호환 가능한 동적 스크립트 경로 탐색
            const possiblePaths = [
              path.join(process.cwd(), 'scripts', 'extract_rag.py'),
              path.join(process.cwd(), 'public', 'extract_rag.py'),
              path.join(process.env.USERPROFILE || 'C:\\Users\\Default', '.gemini', 'ISMP2', 'extract_rag.py'),
              'C:\\Users\\KITC\\.gemini\\ISMP2\\extract_rag.py'
            ];

            const scriptPath = possiblePaths.find(p => fs.existsSync(p)) || possiblePaths[0];
            
            exec(`${pythonCmd} "${scriptPath}"`, (error, stdout, stderr) => {
              if (error) {
                console.error(`Exec error: ${error}`);
                res.statusCode = 500;
                res.setHeader('Content-Type', 'application/json');
                res.end(JSON.stringify({ success: false, error: `스크립트 실행 오류 (${scriptPath}): ${error.message}` }));
                return;
              }
              res.statusCode = 200;
              res.setHeader('Content-Type', 'application/json');
              res.end(JSON.stringify({ success: true, message: 'Indexing completed', scriptUsed: scriptPath }));
            });
          } else if (req.url === '/api/save-rag' && req.method === 'POST') {
            let body = '';
            req.on('data', chunk => { body += chunk.toString(); });
            req.on('end', async () => {
              try {
                const newDoc = JSON.parse(body);
                const fs = await import('fs');
                const path = await import('path');
                const filePath = path.join(process.cwd(), 'public', 'rag_data.json');
                
                let currentData = [];
                if (fs.existsSync(filePath)) {
                  currentData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
                }
                
                // 중복 방지 (파일명 기준)
                const exists = currentData.some(d => d.title === newDoc.title);
                if (!exists) {
                  currentData.push(newDoc);
                  fs.writeFileSync(filePath, JSON.stringify(currentData, null, 2));
                }
                
                res.statusCode = 200;
                res.setHeader('Content-Type', 'application/json');
                res.end(JSON.stringify({ success: true, duplicated: exists }));
              } catch (e) {
                res.statusCode = 500;
                res.end(JSON.stringify({ success: false, error: e.message }));
              }
            });
          } else if (req.url === '/api/delete-rag' && req.method === 'POST') {
            let body = '';
            req.on('data', chunk => { body += chunk.toString(); });
            req.on('end', async () => {
              try {
                const { id } = JSON.parse(body);
                const fs = await import('fs');
                const path = await import('path');
                const filePath = path.join(process.cwd(), 'public', 'rag_data.json');
                
                if (fs.existsSync(filePath)) {
                  let currentData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
                  const updatedData = currentData.filter(d => d.id !== id);
                  fs.writeFileSync(filePath, JSON.stringify(updatedData, null, 2));
                }
                
                res.statusCode = 200;
                res.setHeader('Content-Type', 'application/json');
                res.end(JSON.stringify({ success: true }));
              } catch (e) {
                res.statusCode = 500;
                res.end(JSON.stringify({ success: false, error: e.message }));
              }
            });
          } else if (req.url && req.url.startsWith('/api/g2b') && (req.method === 'POST' || req.method === 'OPTIONS')) {
            // CORS preflight 처리
            if (req.method === 'OPTIONS') {
              res.statusCode = 200;
              res.setHeader('Access-Control-Allow-Origin', '*');
              res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
              res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
              res.end();
              return;
            }

            // 로컬 개발 환경: 조달청 나라장터 API CORS 프록시 (POST JSON, Node.js https 사용)
            try {
              const https = await import('https');

              // POST body 읽기
              let body_str = '';
              await new Promise((resolve) => {
                req.on('data', chunk => { body_str += chunk.toString(); });
                req.on('end', resolve);
              });

              let body_json = {};
              try { body_json = JSON.parse(body_str); } catch (e) {}

              const service_key = body_json.service_key || '';
              const inqry_bgn_dt = body_json.inqry_bgn_dt || '';
              const inqry_end_dt = body_json.inqry_end_dt || '';

              if (!service_key) {
                res.statusCode = 400;
                res.setHeader('Content-Type', 'application/json');
                res.end(JSON.stringify({ error: 'service_key가 필요합니다.' }));
                return;
              }

              // API 키 원본 그대로 사용 (POST body이므로 URL 인코딩 오염 없음)
              // Decoding 키(특수문자 포함)를 URL 파라미터로 넣을 때만 encodeURIComponent 적용
              const key_raw = service_key;
              const key_encoded = encodeURIComponent(service_key);

              // Node.js 내장 https.get 헬퍼
              const https_get = (url) => new Promise((resolve, reject) => {
                https.get(url, (api_res) => {
                  let data = '';
                  api_res.setEncoding('utf8');
                  api_res.on('data', chunk => { data += chunk; });
                  api_res.on('end', () => resolve(data));
                  api_res.on('error', reject);
                }).on('error', reject);
              });

              const service_type = (body_json.service_type || 'prespec').toLowerCase(); // 'prespec' | 'bid' | 'orderplan'
              const key_variants = [...new Set([key_raw, key_encoded])];

              let endpoints = [];
              if (service_type === 'bid') {
                // 입찰공고
                endpoints = [
                  `https://apis.data.go.kr/1230000/ao/PubDataOpnStdService/getDataSetOpnStdBidPblancInfo`,
                  `https://apis.data.go.kr/1230000/BidPublicInfoService02/getBidPblancListInfoServc`,
                  `https://apis.data.go.kr/1230000/BidPublicInfoService02/getBidPblancListInfoServcPPSSrch`
                ];
              } else if (service_type === 'orderplan') {
                // 발주계획
                endpoints = [
                  `https://apis.data.go.kr/1230000/ao/OrderPlanSttusService/getOrderPlanSttusList`,
                  `https://apis.data.go.kr/1230000/OrderPlanSttusService/getOrderPlanSttusList`
                ];
              } else {
                // 사전규격 (기본)
                endpoints = [
                  `https://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoServcPPSSrch`,
                  `https://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoServc`,
                  `https://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoThngPPSSrch`,
                  `https://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoThng`
                ];
              }

              let last_error = '';
              let success = false;

              for (const endpoint of endpoints) {
                for (const key of key_variants) {
                  let api_url = `${endpoint}?serviceKey=${key}&type=json&numOfRows=100&pageNo=1`;
                  if (inqry_bgn_dt) api_url += `&inqryBgnDt=${inqry_bgn_dt}`;
                  if (inqry_end_dt) api_url += `&inqryEndDt=${inqry_end_dt}`;

                  try {
                    console.log(`[G2B 프록시 - ${service_type}] 호출: ${endpoint.split('/').slice(-2).join('/')} (key: ${key.substring(0, 8)}...)`);
                    const text = await https_get(api_url);

                    if (text.includes('SERVICE_KEY_IS_NOT_REGISTERED_ERROR') || text.includes('SERVICE_KEY_IS_NOT_REGISTERED')) {
                      last_error = 'SERVICE_KEY_IS_NOT_REGISTERED';
                      continue;
                    }
                    if (text.includes('LIMITED_NUMBER_OF_SERVICE_REQUESTS')) {
                      last_error = 'LIMITED_NUMBER_OF_SERVICE_REQUESTS';
                      continue;
                    }
                    if (text.includes('NO_OPENAPI_SERVICE_ERROR')) {
                      last_error = `엔드포인트 미지원: ${endpoint}`;
                      continue;
                    }

                    // 정상 데이터 응답 성공
                    res.statusCode = 200;
                    res.setHeader('Content-Type', 'application/json; charset=utf-8');
                    res.setHeader('Access-Control-Allow-Origin', '*');
                    res.end(text);
                    success = true;
                    break;
                  } catch (fe) {
                    last_error = fe.message;
                  }
                }
                if (success) break;
              }

              if (!success) {
                res.statusCode = 200;
                res.setHeader('Content-Type', 'application/json; charset=utf-8');
                res.setHeader('Access-Control-Allow-Origin', '*');
                res.end(JSON.stringify({
                  response: {
                    header: { resultCode: '99', resultMsg: last_error || 'SYNCING' },
                    body: { items: [] }
                  }
                }));
              }
            } catch (err) {
              console.error('[G2B 프록시] 서버 오류:', err);
              res.statusCode = 200;
              res.setHeader('Content-Type', 'application/json; charset=utf-8');
              res.end(JSON.stringify({
                response: {
                  header: { resultCode: '99', resultMsg: err.message },
                  body: { items: [] }
                }
              }));
            }
          } else {
            next();
          }
        });
      }
    }
  ],
  base: './',
})
