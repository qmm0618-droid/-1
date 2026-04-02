#!/usr/bin/env node
/**
 * 课件生成服务 - Vercel 部署版本
 */

const express = require('express');
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const app = express();
app.use(express.json());

function parseOutline(text) {
    const slides = [];
    const lines = text.trim().split('\n');
    let currentSection = null;
    
    for (let line of lines) {
        line = line.trim();
        if (!line) continue;
        
        if (line.startsWith('## ')) {
            currentSection = line.substring(3);
            slides.push({ type: 'section', title: currentSection });
        } else if (line.startsWith('### ')) {
            slides.push({ type: 'subsection', title: line.substring(3), content: '' });
        } else if (line.startsWith('# ')) {
            currentSection = line.substring(2);
            slides.push({ type: 'section', title: currentSection });
        } else if (currentSection) {
            slides.push({ type: 'content', title: currentSection, content: line });
        } else {
            slides.push({ type: 'content', title: '内容', content: line });
        }
    }
    return slides;
}

async function generatePPT(title, outline) {
    const pres = new PptxGenJS();
    pres.layout = 'LAYOUT_16x9';
    
    const slides = parseOutline(outline);
    
    // 封面
    let slide = pres.addSlide();
    slide.addText(title, { x: 0.5, y: 2.5, w: '90%', fontSize: 44, bold: true, color: '363636' });
    slide.addText('课件生成', { x: 0.5, y: 4, w: '90%', fontSize: 24, color: '666666' });
    
    // 内容页
    for (let s of slides) {
        slide = pres.addSlide();
        if (s.type === 'section') {
            slide.addText(s.title, { x: 0.5, y: 2.5, w: '90%', fontSize: 40, bold: true, color: '363636', align: 'center' });
        } else {
            slide.addText(s.title, { x: 0.5, y: 0.5, w: '90%', fontSize: 32, bold: true, color: '363636' });
            if (s.content) {
                slide.addText(s.content, { x: 0.5, y: 1.5, w: '90%', fontSize: 18, color: '363636' });
            }
        }
    }
    
    return await pres.writeToBuffer({ template: null });
}

const htmlPage = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>课件制作工具</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; padding: 20px; }
        .container { max-width: 900px; margin: 0 auto; }
        .header { text-align: center; color: white; margin-bottom: 30px; }
        .header h1 { font-size: 2.5rem; margin-bottom: 10px; }
        .card { background: white; border-radius: 16px; padding: 30px; box-shadow: 0 20px 60px rgba(0,0,0,0.2); margin-bottom: 20px; }
        .form-group { margin-bottom: 20px; }
        .form-group label { display: block; font-weight: 600; margin-bottom: 8px; color: #333; }
        .form-group input, .form-group textarea { width: 100%; padding: 12px 16px; border: 2px solid #e0e0e0; border-radius: 8px; font-size: 16px; }
        .form-group input:focus, .form-group textarea:focus { outline: none; border-color: #667eea; }
        .form-group textarea { min-height: 300px; font-family: 'Monaco', 'Menlo', monospace; line-height: 1.6; resize: vertical; }
        .help-text { font-size: 14px; color: #666; margin-top: 6px; }
        .btn { padding: 14px 32px; border: none; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; transition: all 0.3s; }
        .btn-primary { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; }
        .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 10px 30px rgba(102, 126, 234, 0.4); }
        .btn-group { display: flex; gap: 12px; flex-wrap: wrap; }
        .btn2 { background: #6c757d; color: white; }
        .btn3 { background: #17a2b8; color: white; }
        .status { padding: 12px 20px; border-radius: 8px; margin-top: 20px; display: none; }
        .status.show { display: block; }
        .status.success { background: #d4edda; color: #155724; }
        .status.error { background: #f8d7da; color: #721c24; }
        .status.loading { background: #cce5ff; color: #004085; }
        .download-link { display: inline-block; margin-top: 10px; padding: 10px 20px; background: #28a745; color: white; text-decoration: none; border-radius: 6px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header"><h1>📚 课件制作工具</h1><p>输入标题和大纲，一键生成 PPT 课件</p></div>
        <div class="card">
            <div class="form-group">
                <label>课件标题</label>
                <input type="text" id="title" placeholder="例如：Python 入门教程">
            </div>
            <div class="form-group">
                <label>课件大纲</label>
                <textarea id="outline" placeholder="## 第一章 基础语法
### 变量定义
Python是动态类型语言
### 数据类型
整数、字符串、列表

## 第二章 控制结构
### 条件语句
if 语句的使用"></textarea>
                <p class="help-text">使用 ## 标记章节，### 标记小节</p>
            </div>
            <div class="btn-group">
                <button class="btn btn-primary" onclick="generatePPT()">生成课件</button>
                <button class="btn btn2" onclick="clearForm()">清空</button>
            </div>
            <div class="status" id="status"></div>
        </div>
    </div>
    <script>
        async function generatePPT() {
            const title = document.getElementById('title').value;
            const outline = document.getElementById('outline').value;
            if (!title.trim()) { showStatus('请输入课件标题', 'error'); return; }
            if (!outline.trim()) { showStatus('请输入课件大纲', 'error'); return; }
            showStatus('正在生成课件...', 'loading');
            try {
                const resp = await fetch('/api/generate', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({title, outline})
                });
                const result = await resp.json();
                if (result.success) {
                    showStatus('✅ 课件生成成功！共 ' + result.slides + ' 页', 'success');
                    const link = document.createElement('a');
                    link.href = result.download_url;
                    link.className = 'download-link';
                    link.textContent = '📥 下载 PPT';
                    link.download = result.filename;
                    document.getElementById('status').appendChild(link);
                } else {
                    showStatus('❌ ' + result.message, 'error');
                }
            } catch(e) { showStatus('❌ ' + e.message, 'error'); }
        }
        function showStatus(msg, type) {
            const el = document.getElementById('status');
            el.className = 'status show ' + type;
            el.textContent = msg;
        }
        function clearForm() {
            document.getElementById('title').value = '';
            document.getElementById('outline').value = '';
            document.getElementById('status').className = 'status';
        }
    </script>
</body>
</html>`;

app.get('/', (req, res) => {
    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.send(htmlPage);
});

app.post('/api/generate', async (req, res) => {
    try {
        const { title, outline } = req.body;
        if (!title || !outline) {
            return res.json({ success: false, message: '缺少参数' });
        }
        
        const buffer = await generatePPT(title, outline);
        const filename = title.replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '_') + '_' + Date.now() + '.pptx';
        
        res.json({
            success: true,
            slides: parseOutline(outline).length + 1,
            filename: filename,
            download_url: `/api/download/${encodeURIComponent(filename)}`
        });
        
        // 保存到临时存储（Vercel 无持久存储，这里返回 base64）
        global.pptBuffers = global.pptBuffers || {};
        global.pptBuffers[filename] = buffer;
        
    } catch (e) {
        res.json({ success: false, message: e.message });
    }
});

app.get('/api/download/:filename', (req, res) => {
    const filename = decodeURIComponent(req.params.filename);
    const buffer = global.pptBuffers?.[filename];
    
    if (buffer) {
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
        res.setHeader('Content-Disposition', 'attachment; filename="' + filename + '"');
        res.send(buffer);
        delete global.pptBuffers[filename];
    } else {
        res.status(404).send('Not found');
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 服务已启动: http://localhost:${PORT}`);
});

module.exports = app;