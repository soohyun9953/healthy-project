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
          } else {
            next();
          }
        });
      }
    }
  ],
  base: './',
})
