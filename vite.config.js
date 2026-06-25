import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import fs from 'fs'
import path from 'path'

const DATA_FILE = path.resolve('./data/sessions.json')
const DATA_DIR  = path.dirname(DATA_FILE)

export default defineConfig({
  plugins: [
    vue(),
    {
      name: 'sessions-api',
      configureServer(server) {
        server.middlewares.use('/api/sessions', (req, res) => {
          res.setHeader('Content-Type', 'application/json; charset=utf-8')
          res.setHeader('Access-Control-Allow-Origin', '*')
          res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
          res.setHeader('Access-Control-Allow-Headers', 'Content-Type')

          if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return }

          if (req.method === 'GET') {
            try {
              const data = fs.existsSync(DATA_FILE)
                ? fs.readFileSync(DATA_FILE, 'utf-8')
                : '{}'
              res.writeHead(200)
              res.end(data)
            } catch (e) {
              res.writeHead(500)
              res.end(JSON.stringify({ error: e.message }))
            }
            return
          }

          if (req.method === 'POST') {
            let body = ''
            req.on('data', chunk => { body += chunk })
            req.on('end', () => {
              try {
                JSON.parse(body)                                  // validate JSON
                if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true })
                fs.writeFileSync(DATA_FILE, body, 'utf-8')
                res.writeHead(200)
                res.end('{"ok":true}')
              } catch (e) {
                res.writeHead(400)
                res.end(JSON.stringify({ error: e.message }))
              }
            })
            return
          }

          res.writeHead(405); res.end()
        })
      },
    },
  ],
})
