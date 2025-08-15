import logging
import os
import sys
import platform
import ctypes

import webview
import darkdetect

from .backend_api import Api
import app.backend_api as backend_api_module
from .utils import image_to_base64, set_dark_title_bar

# --- 前端HTML模板 ---
# 再次修正，并对脚本部分进行了格式化以提高可读性
html_template = """
<!DOCTYPE html><html lang="zh-CN"><head>
<meta charset="UTF-8"><title>Markdown to Word Converter by Youkies</title>
<link rel="stylesheet" href="static/css/easymde.min.css">
<link rel="stylesheet" href="static/css/classic.min.css"/>
<style>
:root {
    --bg-color: #f5f5f5; --editor-bg: #ffffff; --text-color: #242424; --border-color: #e0e0e0;
    --header-bg: #f5f5f5; --button-bg: #e0e0e0; --button-hover-bg: #d0d0d0;
    --info-text: #555555; --scrollbar-thumb: #cdcdcd; --scrollbar-thumb-hover: #a8a8a8;
    --danger-button-bg: #ffcdd2; --danger-button-hover-bg: #ef9a9a;
    --loader-color: #0078d4; --input-bg: #ffffff;
    --preview-bg: #ffffff; --preview-text: #212529; --preview-border: #dee2e6; 
    --preview-code-bg: #f8f9fa; --preview-link: #0d6efd; 
    --preview-blockquote-text: #6c757d; --preview-blockquote-border: #e9ecef;
}
@media (prefers-color-scheme: dark) {
    :root {
        --bg-color: #252526; --editor-bg: #252526; --text-color: #d4d4d4;
        --border-color: #3c3c3c; --header-bg: #252526; --button-bg: #3e3e40;
        --button-hover-bg: #5a5a5c; --info-text: #a0a0a0; --scrollbar-thumb: #4a4a4a;
        --danger-button-bg: #5c4343; --danger-button-hover-bg: #6e5151;
        --scrollbar-thumb-hover: #555555; --loader-color: #61afef; --input-bg: #3c3c3c;
        --preview-bg: #252526; --preview-text: #d4d4d4; --preview-border: #4a4a4a; 
        --preview-code-bg: #1e1e1e; --preview-link: #61afef; 
    }
    /* 暗色模式下的预加载样式 */
    html, body {
        background-color: #252526 !important;
    }
}
html {
    height: 100%;
    visibility: visible !important;
    background-color: var(--bg-color);
}
body, html { 
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif; 
    margin: 0; 
    padding: 0; 
    height: 100%; 
    overflow: hidden; 
    background-color: var(--bg-color);
    color: var(--text-color); 
}
/* 预加载样式，确保窗口创建时不出现白色闪烁 */
.app-container {
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    visibility: hidden;
    opacity: 0;
    background-color: var(--bg-color);
    transition: opacity 0.2s;
}
#loader {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    height: 100%;
    width: 100%;
    position: fixed;
    top: 0;
    left: 0;
    background-color: var(--bg-color);
    color: var(--text-color);
    z-index: 1000;
    transition: opacity 0.3s;
}
.spinner {
    width: 40px;
    height: 40px;
    border: 4px solid rgba(0,0,0,0.1);
    border-left-color: var(--loader-color);
    border-radius: 50%;
    animation: spin 1s linear infinite;
}
@keyframes spin {
    to { transform: rotate(360deg); }
}
@media (prefers-color-scheme: dark) {
    .spinner {
        border: 4px solid rgba(255,255,255,0.1);
        border-left-color: var(--loader-color);
    }
}
.app-header {
    display: flex;
    justify-content: flex-end;
    align-items: center;
    padding: 8px 12px;
    background-color: var(--header-bg);
    border-bottom: 1px solid var(--border-color);
    flex-shrink: 0;
    gap: 8px;
    height: 40px;
}
.header-group {
    display: flex;
    align-items: center;
    gap: 8px;
    flex-shrink: 0;
}
button {
    background-color: var(--button-bg);
    color: var(--text-color);
    border: 1px solid var(--border-color);
    padding: 0 14px;
    border-radius: 5px;
    cursor: pointer;
    font-size: 13px;
    transition: all 0.15s ease-in-out;
    white-space: nowrap;
    height: 38px;
    line-height: 38px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
}
button:hover { 
    background-color: var(--button-hover-bg);
    border-color: var(--button-hover-bg);
}
button:active {
    transform: translateY(1px);
}
button:disabled { 
    background-color: var(--info-text); 
    cursor: not-allowed; 
    transform: none;
    box-shadow: none;
    border-color: transparent;
}
#clear-btn { background-color: var(--danger-button-bg); }
#clear-btn:hover { 
    background-color: var(--danger-button-hover-bg); 
    border-color: var(--danger-button-hover-bg);
}
.action-btn { background-color: var(--button-bg); }
.cancel-btn { background-color: #6c757d !important; }
.cancel-btn:hover { 
    background-color: #5a6268 !important; 
    border-color: #5a6268 !important;
}
.export-controls {
    flex-grow: 0.4;
    display: flex;
    align-items: center;
    gap: 6px;
    min-width: 0;
    max-width: 38%;
}
#folder-path, #file-name-input {
    background-color: var(--input-bg);
    border: 1px solid var(--border-color);
    color: var(--text-color);
    padding: 0 10px;
    border-radius: 5px;
    font-size: 13px;
    height: 38px;
    box-sizing: border-box;
}
#folder-path {
    flex-grow: 1;
    min-width: 40px;
    max-width: 150px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    cursor: default;
    line-height: 36px;
}
#file-name-input {
    width: 180px;
}
.input-label {
    font-size: 13px;
    color: var(--info-text);
    white-space: nowrap;
    max-width: 100px;
    overflow: hidden;
    text-overflow: ellipsis;
    display: inline-block;
    vertical-align: middle;
}
.template-info {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 13px; 
    flex-shrink: 0; 
    min-width: 0;
}
#template-status {
    color: var(--info-text);
    font-size: 13px;
    max-width: 100px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    display: inline-block;
    vertical-align: middle;
}
#style-settings-btn, #template-presets-btn {
    padding: 0 14px;
    min-width: 80px;
}
.header-group:last-child {
    margin-left: auto;
}
/* 固定最后两个按钮的布局 */
.header-group:nth-last-child(1) {
    min-width: 240px;
    justify-content: flex-end;
}
/* 确保模板状态文本在空间足够时能显示更多内容 */
@media (min-width: 1300px) {
    #template-status {
        max-width: 150px;
    }
}
.main-content{display:flex;flex-grow:1;height:calc(100vh - 57px - 23px);}
.editor-pane{flex:1;display:flex;flex-direction:column;border-right:1px solid var(--border-color);min-width:300px;position: relative;}
.preview-pane{flex:1; min-width:300px; position: relative; display: flex;}
#preview-div {flex-grow: 1; overflow-y: auto; padding: 25px; background-color:var(--preview-bg); color: var(--preview-text); font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif; line-height:1.6;}
.preview-pane a { color: var(--preview-link); } 
.preview-pane table { border-collapse:collapse; width:100%; margin:1em 0; }
.preview-pane th, .preview-pane td { border:1px solid var(--preview-border); padding:8px 12px; text-align:left; } 
.preview-pane th { background-color: var(--preview-code-bg); }
.preview-pane code { background-color: rgba(128,128,128,0.1); padding:.2em .4em; border-radius:4px; font-size:85%; font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, Courier, monospace; }
.preview-pane pre { background-color: var(--preview-code-bg); padding:1em; border-radius:5px; overflow-x:auto; }
.preview-pane pre code { background-color: transparent; padding: 0; border-radius: 0; font-size: inherit;}
.preview-pane blockquote { border-left:.25em solid var(--preview-blockquote-border); padding:0 1em; color: var(--preview-blockquote-text); margin-left:0; }
.preview-pane img { max-width:100%; height:auto; }
.floating-action-group { position: absolute; bottom: 55px; right: 25px; z-index: 100; display: flex; flex-direction: column; gap: 12px; opacity: 0.8; transition: opacity 0.2s; }
.floating-action-group:hover { opacity: 1; }
.floating-action-group button { box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
@media (prefers-color-scheme: dark) { .floating-action-group button { box-shadow: 0 2px 8px rgba(0,0,0,0.25); } }
button img {
    width: 16px;
    height: 16px;
    margin-left: 8px;
    vertical-align: middle;
}
.EasyMDEContainer{height:100%;display:flex;flex-direction:column;border:none;}
.EasyMDEContainer .CodeMirror{flex-grow:1;background-color:var(--editor-bg);color:var(--text-color);border:none!important;}
.EasyMDEContainer .editor-toolbar{display:none;}
.editor-statusbar{background-color:var(--header-bg);color:var(--text-color);border:none;}
.CodeMirror-selected { background: #717171 !important; }
::-webkit-scrollbar{width:10px;height:10px;} ::-webkit-scrollbar-track{background:var(--bg-color);} ::-webkit-scrollbar-thumb{background-color:var(--scrollbar-thumb);border-radius:5px;} ::-webkit-scrollbar-thumb:hover{background-color:var(--scrollbar-thumb-hover);}
.notification{position:fixed;bottom:40px;left:50%;transform:translateX(-50%);padding:10px 20px;border-radius:5px;color:white;z-index:1000;opacity:0;transition:opacity .5s;}
.notification.show{opacity:1;} .notification.success{background-color:#28a745;} .notification.error{background-color:#dc3545;} .notification.info{background-color:#17a2b8;}
.app-footer { padding: 4px 15px; background-color: var(--header-bg); border-top: 1px solid var(--border-color); color: var(--info-text); font-size: 0.8em; display: flex; justify-content: space-between; align-items: center; flex-shrink: 0; height: 14px; }
#about-btn { cursor: pointer; text-decoration: underline; } #about-btn:hover { color: var(--button-bg); }
.modal-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background-color: rgba(0, 0, 0, 0.6); display: none; justify-content: center; align-items: center; z-index: 1001; opacity: 0; transition: opacity 0.3s ease; }
.modal-overlay.show { display: flex; opacity: 1; }
.modal-content { background-color: var(--editor-bg); padding: 25px; border-radius: 8px; box-shadow: 0 5px 15px rgba(0, 0, 0, 0.3); width: 90%; max-width: 650px; text-align: left; }
.modal-content h3 { margin-top: 0; color: var(--text-color); text-align: center; } .modal-content p { color: var(--info-text); word-wrap: break-word; }
#export-path-display { font-style: italic; font-size: 0.9em; }
.modal-buttons { margin-top: 20px; display: flex; justify-content: flex-end; flex-wrap: wrap; gap: 10px; }
.social-qr-codes { display: flex; justify-content: center; gap: 30px; margin-top: 20px; }
.qr-code-item { text-align: center; } .qr-code-item img { width: 180px; height: auto; border-radius: 8px; border: 1px solid var(--border-color); }
.qr-code-item p { margin-top: 5px; font-size: 0.9em; }
#style-settings-dialog-overlay .modal-content { max-width: 500px; }
.style-section { margin-bottom: 15px; padding-bottom: 15px; border-bottom: 1px solid var(--border-color); }
.style-section:last-of-type { border-bottom: none; margin-bottom: 0; }
.style-section h4 { margin-top: 0; margin-bottom: 15px; color: var(--text-color); }
.style-grid { display: grid; grid-template-columns: 50px 1fr; gap: 10px 15px; align-items: center; }
.style-grid label { font-size: 14px; color: var(--info-text); text-align: right; }
.style-grid select, .style-grid input, .style-grid .size-controls { width: 100%; }
.size-controls { display: flex; gap: 10px; }
.size-controls > * { flex: 1; }
.color-picker-btn-wrapper { position: relative; width: 38px; height: 38px; }
.pcr-button { width: 100%; height: 100%; border-radius: 5px; border: 1px solid var(--border-color) !important; }
.pcr-app { z-index: 1002; }
select, input[type="number"] { -webkit-appearance: none; -moz-appearance: none; appearance: none; background-color: var(--input-bg); border: 1px solid var(--border-color); border-radius: 5px; padding: 7px 10px; font-size: 13px; color: var(--text-color); box-sizing: border-box; }
select { background-image: url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'%3e%3cpath fill='none' stroke='%236c757d' stroke-linecap='round' stroke-linejoin='round' stroke-width='2' d='M2 5l6 6 6-6'/%3e%3c/svg%3e"); background-repeat: no-repeat; background-position: right 0.75rem center; background-size: 16px 12px; }
@media (prefers-color-scheme: dark){ select { background-image: url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'%3e%3cpath fill='none' stroke='%23a0a0a0' stroke-linecap='round' stroke-linejoin='round' stroke-width='2' d='M2 5l6 6 6-6'/%3e%3c/svg%3e"); } }
input[type="number"] { -moz-appearance: textfield; }
input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
#template-dialog-overlay .modal-content { max-width: 700px; }
.template-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; }
.template-card { display: flex; flex-direction: column; align-items: center; justify-content: center; background-color: var(--input-bg); border: 1px solid var(--border-color); border-radius: 8px; padding: 20px; text-align: center; cursor: pointer; transition: all 0.2s ease-in-out; position: relative; }
.template-card.selected { border-color: var(--button-bg); box-shadow: 0 0 0 2px var(--button-bg); }
.template-card:hover { transform: translateY(-5px); box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
.template-card svg { width: 48px; height: 48px; margin-bottom: 15px; color: var(--button-bg); }
.template-card span { font-size: 14px; color: var(--text-color); }
.template-card .checkmark { display: none; position: absolute; bottom: 8px; right: 8px; width: 20px; height: 20px; background-color: #28a745; color: white; border-radius: 50%; font-size: 14px; line-height: 20px; text-align: center; }
.template-card.selected .checkmark { display: block; }
.custom-tooltip { position: fixed; display: none; background-color: var(--header-bg); color: var(--text-color); border: 1px solid var(--border-color); border-radius: 5px; padding: 10px; z-index: 1005; box-shadow: 0 4px 8px rgba(0,0,0,0.2); font-size: 12px; pointer-events: none; line-height: 1.6; }
.tooltip-line { display: flex; align-items: center; gap: 8px; }
.tooltip-color-swatch { display: inline-block; width: 12px; height: 12px; border: 1px solid var(--border-color); border-radius: 2px; flex-shrink: 0; }
.toolbar-separator {
    width: 1px;
    height: 24px;
    background-color: var(--border-color);
    margin: 0 8px;
}
</style>
</head><body>
<script src="static/js/easymde.min.js"></script>
<script src="static/js/marked.min.js"></script>
<script src="static/js/pickr.min.js"></script>
<script>
MathJax = {
  tex: { 
    inlineMath: [['$', '$'], ['\\(', '\\)']], 
    displayMath: [['$$', '$$'], ['\\[', '\\]']], 
    processEscapes: true 
  },
  chtml: {
    fontURL: 'https://cdn.jsdelivr.net/npm/mathjax@3/es5/output/chtml/fonts/woff-v2'
  }
};
</script>
<script type="text/javascript" id="MathJax-script" async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>
<div id="loader"><div class="spinner"></div><p>正在加载...</p></div>
<div class="app-container">
    <header class="app-header">
        <button id="select-folder-btn">选择输出位置</button>
        <label class="input-label" for="folder-path">当前输出位置：</label>
        <span id="folder-path" title="当前保存路径"></span>
        <span class="toolbar-separator"></span>
        <button id="save-file-btn">保存为 Word</button>
        <label for="file-name-input" class="input-label">输出文件名：</label>
        <input type="text" id="file-name-input" placeholder="文件名">
        <span class="toolbar-separator"></span>
        <button id="style-settings-btn">样式设置</button>
        <button id="template-presets-btn">预设模板</button>
        <span id="template-status" title="当前模板">(默认)</span>
    </header>
    <main class="main-content">
        <div class="editor-pane"><textarea id="markdown-editor">## 🚀 欢迎使用 Markdown to Word 转换器 v3.0

感谢您选择本工具！新版本，新体验，希望能更好地帮助到你！

> **快速上手**：在左侧窗格中粘贴 Markdown 文本（例如 AI 回答），右侧会实时展示预览效果（仅供参考，实际以输出文档为准）。

### ✨ 核心功能

#### 1. 数学公式支持
本工具全面兼容 LaTeX 数学公式，无论是行内公式还是复杂的块级公式，都能精准渲染。

* **行内公式**: `$a^2+b^2=c^2$` 或 `\\(a^2+b^2=c^2\\)`
* **块级公式**: `$$...$$` 或 `\\[...\\]`
    * 示例:
        \\[
        \\mu = \\frac{1}{t} \\ln \\frac{OD_{\\text{final}}}{OD_{\\text{initial}}}
        \\]

> 如果您输入的公式有误，转换时右下角将会弹出错误提示。

#### 2. 支持表格创建
使用标准的 Markdown 语法，轻松创建和对齐表格。

| 表头1 | 表头2 | 表头3 |
| :--- | :---: | ---: |
| 居左 | 居中 | 居右 |
| 单元格 | 单元格 | 单元格 |

### 🌟 v3.0 全新升级

* **界面焕新**：对界面进行了彻底的重新设计，调整了功能键的位置与颜色，整体视觉更和谐，操作更直观。
* **一键复制到 Word**：支持"复制到 Word"功能，可一键将预览区的所有内容（包括格式、表格、公式）直接粘贴到 Microsoft Word 或 WPS 中。
    > **重要提示**:
    > 由于 WPS 和 Microsoft Word 文件格式存在差异，请选择对应选项（Microsoft Word/WPS），否则可能导致数学公式无法正常显示。
* **框架重构与智能记忆**：重构了底层应用框架，运行更稳定流畅。同时，您的导出路径、模板选择和自定义样式等偏好设置将被自动保存，重启后无需重新配置。
* **文本处理**：新增文本处理设置功能，目前只添加了自动删除分隔线选项。

---

**作者**: (酷安@Youkies)
**版本**: 3.0

如果您觉得这款工具对您有帮助，不妨**请作者喝一杯咖啡**！您的支持是持续更新的最大动力。
</textarea>
            <div class="floating-action-group">
                <button id="open-file-btn" title="打开文件">打开文件<img src="static/icons/Markdown.png" alt="Markdown"></button>
                <button id="text-processing-btn" title="文本处理设置">文本处理<img src="static/icons/edit.png" alt="Edit"></button>
                <button id="paste-btn" title="从剪贴板粘贴内容">一键粘贴<img src="static/icons/Paste.png" alt="Paste"></button>
                <button id="clear-btn" title="清空所有内容">一键清空<img src="static/icons/Clear.png" alt="Clear"></button>
            </div>
        </div>
        <div class="preview-pane">
            <div id="preview-div"></div>
            <div class="floating-action-group">
                <button id="copy-to-word-btn" title="通过后台 Word 带格式复制内容">复制到 Word<img src="static/icons/Microsoft Word.png" alt="Word"></button>
                <button id="copy-to-wps-btn" title="通过后台 WPS 带格式复制内容">复制到 WPS<img src="static/icons/wps.png" alt="WPS"></button>
            </div>
        </div>
    </main>
    <footer class="app-footer">
        <span>Markdown to Word Converter by Youkies</span>
        <div><span id="about-btn">请作者喝一杯咖啡</span><span style="margin: 0 5px;">|</span><span>Version 3.0</span></div>
    </footer>
</div>
<div id="notification-box" class="notification"></div>
<div id="export-success-dialog-overlay" class="modal-overlay">
    <div class="modal-content" style="text-align: center;">
        <h3>导出成功！</h3><p>文件已保存至：<br><span id="export-path-display"></span></p>
        <div class="modal-buttons"><button id="open-file-btn-success" class="action-btn">打开文件</button><button id="open-folder-btn-success" class="action-btn">打开文件夹</button><button class="close-btn cancel-btn">关闭</button></div>
    </div>
</div>
<div id="about-dialog-overlay" class="modal-overlay">
    <div class="modal-content" style="text-align: center;">
        <h3>关于本软件</h3><p>这是一款简洁高效的 Markdown 到 Word 转换工具，支持实时预览、自定义 Word 模板和所有标准 LaTeX 数学公式。</p><p>酷安@Youkies</p><p>谢谢你的赞助！</p>
        <div class="social-qr-codes">
            <div class="qr-code-item"><img src="__WECHAT_QR_PLACEHOLDER__" alt="微信二维码"><p>微信</p></div>
            <div class="qr-code-item"><img src="__ALIPAY_QR_PLACEHOLDER__" alt="支付宝二维码"><p>支付宝</p></div>
        </div>
        <div class="modal-buttons"><button class="close-btn cancel-btn">关闭</button></div>
    </div>
</div>
<div id="style-settings-dialog-overlay" class="modal-overlay">
    <div class="modal-content">
        <h3>样式设置</h3>
        <div class="style-section" data-style-for="body">
            <h4>正文设置</h4>
            <div class="style-grid">
                <label for="body-font">字体</label><select id="body-font"></select>
                <label for="body-size">字号</label><div class="size-controls"><input id="body-size" type="number" step="0.5"><select id="body-size-cn"></select></div>
                <label for="body-color">颜色</label><div id="body-color-picker" class="color-picker-btn-wrapper"></div>
            </div>
        </div>
        <div class="style-section" data-style-for="h1">
            <h4>一级标题 (H1)</h4>
            <div class="style-grid">
                <label for="h1-font">字体</label><select id="h1-font"></select>
                <label for="h1-size">字号</label><div class="size-controls"><input id="h1-size" type="number" step="1"><select id="h1-size-cn"></select></div>
                <label for="h1-color">颜色</label><div id="h1-color-picker" class="color-picker-btn-wrapper"></div>
            </div>
        </div>
        <div class="style-section" data-style-for="h2">
            <h4>二级标题 (H2)</h4>
            <div class="style-grid">
                <label for="h2-font">字体</label><select id="h2-font"></select>
                <label for="h2-size">字号</label><div class="size-controls"><input id="h2-size" type="number" step="1"><select id="h2-size-cn"></select></div>
                <label for="h2-color">颜色</label><div id="h2-color-picker" class="color-picker-btn-wrapper"></div>
            </div>
        </div>
        <div class="style-section" data-style-for="h3">
            <h4>三级标题 (H3)</h4>
            <div class="style-grid">
                <label for="h3-font">字体</label><select id="h3-font"></select>
                <label for="h3-size">字号</label><div class="size-controls"><input id="h3-size" type="number" step="1"><select id="h3-size-cn"></select></div>
                <label for="h3-color">颜色</label><div id="h3-color-picker" class="color-picker-btn-wrapper"></div>
            </div>
        </div>
        <div class="modal-buttons">
            <button id="save-style-settings-btn" class="action-btn">保存并关闭</button>
            <button id="cancel-style-settings-btn" class="cancel-btn">取消</button>
        </div>
    </div>
</div>
<div id="template-dialog-overlay" class="modal-overlay">
    <div class="modal-content">
        <h3>选择一个模板</h3>
        <div class="template-grid">
            <div class="template-card" data-preset="general">
                <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
                <span>常规文档</span><div class="checkmark">✔</div>
            </div>
            <div class="template-card" data-preset="academic">
                <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path></svg>
                <span>学术论文</span><div class="checkmark">✔</div>
            </div>
            <div class="template-card" data-preset="business">
                 <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="6" y1="2" x2="6" y2="8"></line><line x1="10" y1="2" x2="10" y2="8"></line><rect x="2" y="14" width="20" height="8" rx="2"></rect><path d="M4 14a2 2 0 0 0-2 2v2a2 2 0 0 0 2 2"></path><path d="M20 14a2 2 0 0 1 2 2v2a2 2 0 0 1-2 2"></path></svg>
                <span>商务报告</span><div class="checkmark">✔</div>
            </div>
             <div class="template-card" data-preset="technical">
                <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="16 18 22 12 16 6"></polyline><polyline points="8 6 2 12 8 18"></polyline></svg>
                <span>技术文档</span><div class="checkmark">✔</div>
            </div>
            <div class="template-card" data-preset="teaching">
                <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2L2 7l10 5 10-5-10-5z"></path><path d="M2 17l10 5 10-5"></path><path d="M2 12l10 5 10-5"></path></svg>
                <span>教学材料</span><div class="checkmark">✔</div>
            </div>
            <div class="template-card" data-preset="government">
                <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2.69l5.66 5.66a8 8 0 1 1-11.31 0z"></path></svg>
                <span>政府公文</span><div class="checkmark">✔</div>
            </div>
            <div class="template-card" data-preset="modern">
                <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect></svg>
                <span>简约现代</span><div class="checkmark">✔</div>
            </div>
            <div class="template-card" data-preset="custom">
                 <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06-.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg>
                <span>自定义...</span><div class="checkmark">✔</div>
            </div>
        </div>
        <div class="modal-buttons">
            <button id="apply-template-btn" class="action-btn">应用</button>
            <button id="cancel-template-btn" class="cancel-btn">取消</button>
        </div>
    </div>
</div>
<div id="template-tooltip" class="custom-tooltip"></div>
<div id="text-processing-dialog-overlay" class="modal-overlay">
    <div class="modal-content">
        <h3>文本处理设置</h3>
        <div class="text-processing-options">
            <label class="switch-label">
                <div class="checkbox-row">
                    <input type="checkbox" id="remove-separators-checkbox">
                    <span class="switch-text">总是删除分隔线</span>
                </div>
                <span class="switch-description">自动删除文本中的所有"---"分隔线</span>
            </label>
        </div>
        <div class="modal-buttons">
            <button id="save-text-processing-btn" class="action-btn">保存并关闭</button>
            <button id="cancel-text-processing-btn" class="cancel-btn">取消</button>
        </div>
    </div>
</div>
<script>
async function initializeApp() {
    try {
        window.app = {
            easyMDE: new EasyMDE({ element: document.getElementById('markdown-editor'), spellChecker: false, status: false, toolbar: false }),
            folderPathEl: document.getElementById('folder-path'),
            fileNameInputEl: document.getElementById('file-name-input'),
            previewDiv: document.getElementById('preview-div'),
            templateStatus: document.getElementById('template-status'),
            styles: {},
            colorPickers: {},
            lastPreset: 'general',

            async renderPreview() {
                const mdText = this.easyMDE.value();
                this.previewDiv.innerHTML = marked.parse(mdText, { gfm: true, breaks: true });
                if (window.MathJax && window.MathJax.typesetPromise) {
                    try { await window.MathJax.typesetPromise([this.previewDiv]); } 
                    catch(err) { console.error('MathJax Typesetting Error:', err); }
                }
            },
            async updateFilename() {
                const content = this.easyMDE.value();
                const filename = await window.pywebview.api.update_filename(content);
                this.fileNameInputEl.value = filename;
            },
            onFileOpened(content) { 
                this.easyMDE.value(content); 
                if(this.easyMDE.codemirror) { this.easyMDE.codemirror.clearHistory(); }
                this.updateFilename(); 
                this.renderPreview(); 
            },
            onTemplateSelected(path) { 
                const baseName = path.split(/[\\\\/]/).pop();
                this.templateStatus.textContent = baseName;
                this.templateStatus.title = path; 
                this.lastPreset = 'custom';
                this.showNotification('已选择自定义模板: ' + baseName, 'success'); 
                document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                document.querySelector('.template-card[data-preset="custom"]').classList.add('selected');
            },
            showNotification(message, type = 'info') {
                const box = document.getElementById('notification-box');
                box.textContent = message; box.className = "notification show " + type;
                setTimeout(() => { box.className = "notification"; }, 4000);
            },
            showExportSuccessDialog(filePath) {
                this.currentFilePath = filePath;
                document.getElementById('export-path-display').textContent = filePath;
                document.getElementById('export-success-dialog-overlay').classList.add('show');
            },
            currentFilePath: '',
            loadStylesToForm(styles) {
                this.styles = JSON.parse(JSON.stringify(styles));
                for (const el in this.styles) {
                    const styleData = this.styles[el];
                    if(!styleData) continue;
                    document.getElementById(`${el}-font`).value = styleData.font;
                    document.getElementById(`${el}-size`).value = styleData.size;
                    updateCnSizeSelect(el, styleData.size);
                    if(this.colorPickers[el]) { this.colorPickers[el].setColor('#' + styleData.color); }
                }
            },
            readStylesFromForm() {
                const newStyles = {};
                const elements = ['body', 'h1', 'h2', 'h3'];
                elements.forEach(el => {
                    newStyles[el] = {
                        font: document.getElementById(`${el}-font`).value,
                        size: document.getElementById(`${el}-size`).value,
                        color: this.colorPickers[el].getColor().toHEXA().toString(0).substring(1)
                    };
                });
                this.styles = newStyles;
                return newStyles;
            },
            async showTextProcessingDialog() {
                try {
                    const settings = await window.pywebview.api.get_text_processing_settings();
                    
                    const removeSeparators = document.getElementById('remove-separators-checkbox');
                    if (settings && settings.remove_separators !== undefined) {
                        removeSeparators.checked = settings.remove_separators;
                    } else {
                        removeSeparators.checked = false;
                    }
                    
                    document.getElementById('text-processing-dialog-overlay').classList.add('show');
                } catch (error) {
                    this.showNotification('加载文本处理设置失败: ' + error, 'error');
                }
            },
            async saveTextProcessingSettings() {
                try {
                    const settings = {
                        remove_separators: document.getElementById('remove-separators-checkbox').checked
                    };
                    
                    const result = await window.pywebview.api.save_text_processing_settings(settings);
                    if (result.success) {
                        this.showNotification('文本处理设置已保存', 'success');
                        
                        const currentText = this.easyMDE.value();
                        if (currentText) {
                            const processedText = await window.pywebview.api.process_text(currentText);
                            if (processedText !== currentText) {
                                this.easyMDE.value(processedText);
                                this.renderPreview();
                            }
                        }
                    } else {
                        this.showNotification('保存文本处理设置失败: ' + result.error, 'error');
                    }
                    
                    document.getElementById('text-processing-dialog-overlay').classList.remove('show');
                } catch (error) {
                    this.showNotification('保存文本处理设置失败: ' + error, 'error');
                }
            },
            async applyTextProcessing() {
                try {
                    const currentText = this.easyMDE.value();
                    if (currentText) {
                        const processedText = await window.pywebview.api.process_text(currentText);
                        if (processedText !== currentText) {
                            this.easyMDE.value(processedText);
                            this.renderPreview();
                            this.showNotification('文本处理已完成', 'success');
                        } else {
                            this.showNotification('没有需要处理的内容', 'info');
                        }
                    }
                } catch (error) {
                    this.showNotification('文本处理失败: ' + error, 'error');
                }
            }
        };

        const cnSizeMap = { '初号': 42, '小初': 36, '一号': 26, '小一': 24, '二号': 22, '小二': 18, '三号': 16, '小三': 15, '四号': 14, '小四': 12, '五号': 10.5, '小五': 9, '六号': 7.5, '小六': 6.5, '七号': 5.5, '八号': 5 };
        const cnSizeOptionsHtml = '<option value="custom">自定义</option>' + Object.keys(cnSizeMap).map(name => `<option value="${cnSizeMap[name]}">${name} (${cnSizeMap[name]}pt)</option>`).join('');
        function updateCnSizeSelect(el, ptValue) {
            const cnSelect = document.getElementById(`${el}-size-cn`);
            const floatPt = parseFloat(ptValue);
            let match = false;
            for (const name in cnSizeMap) { if (cnSizeMap[name] === floatPt) { cnSelect.value = floatPt; match = true; break; } }
            if (!match) cnSelect.value = 'custom';
        }
        const styleElements = ['body', 'h1', 'h2', 'h3'];
        styleElements.forEach(el => {
            const ptInput = document.getElementById(`${el}-size`);
            const cnSelect = document.getElementById(`${el}-size-cn`);
            cnSelect.innerHTML = cnSizeOptionsHtml;
            ptInput.addEventListener('input', () => updateCnSizeSelect(el, ptInput.value));
            cnSelect.addEventListener('change', () => { if (cnSelect.value !== 'custom') ptInput.value = cnSelect.value; });
        });

        const [initialInfo, systemFonts, presetStyles] = await Promise.all([
            window.pywebview.api.get_initial_info(),
            window.pywebview.api.get_system_fonts(),
            window.pywebview.api.get_preset_styles()
        ]);

        window.app.styles = initialInfo.styles;
        window.app.lastPreset = initialInfo.last_preset || 'general';
        const fontSelects = document.querySelectorAll('select[id$="-font"]');
        const fontOptionsHtml = systemFonts.map(font => `<option value="${font}">${font}</option>`).join('');
        fontSelects.forEach(select => select.innerHTML = fontOptionsHtml);

        styleElements.forEach(el => {
            const picker = Pickr.create({
                el: `#${el}-color-picker`,
                theme: 'classic',
                default: `#${window.app.styles[el]?.color || '000000'}`,
                swatches: [
                    '#FFFFFF', '#000000', '#808080', '#1F497D', '#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646', 
                    '#F2F2F2', '#7F7F7F', '#D9D9D9', '#C6D9F1', '#DCE6F2', '#F2DCDB', '#EBF1DE', '#E5E0EC', '#DBEEF4', '#FDE9D9', 
                    '#D8D8D8', '#595959', '#BFBFBF', '#8DB3E2', '#B8CCE4', '#E6B9B8', '#D7E4BD', '#CCC1D9', '#B7DEE8', '#FBD5B5', 
                    '#BFBFBF', '#3F3F3F', '#9F9F9F', '#548DD4', '#95B3D7', '#D99694', '#C3D69B', '#B2A2C7', '#93CDDD', '#FAC08F', 
                    '#A5A5A5', '#262626', '#7F7F7F', '#2D65A2', '#7793B9', '#C00000', '#92B434', '#843C0C', '#538DD5', '#00B050'
                ],
                components: {
                    preview: true, opacity: false, hue: true,
                    interaction: { hex: true, rgba: true, input: true, save: true, clear: false }
                },
                i18n: { 'btn:save': '确定', 'btn:clear': '清除' }
            });
            window.app.colorPickers[el] = picker;
        });

        window.app.folderPathEl.textContent = initialInfo.export_directory;
        window.app.folderPathEl.title = initialInfo.export_directory;
        window.app.fileNameInputEl.value = "无标题";

        if (initialInfo.last_preset === 'custom') {
            if (initialInfo.template_path) {
                const baseName = initialInfo.template_path.split(/[\\\\/]/).pop();
                window.app.templateStatus.textContent = baseName;
                window.app.templateStatus.title = initialInfo.template_path;
            } else {
                window.app.templateStatus.textContent = '自定义样式';
                window.app.templateStatus.title = '自定义样式';
            }
        } else {
            const preset = initialInfo.last_preset || 'general';
            const card = document.querySelector(`.template-card[data-preset="${preset}"]`);
            if (card) {
                const cardText = card.querySelector('span').textContent.trim();
                window.app.templateStatus.textContent = cardText;
                window.app.templateStatus.title = `预设: ${cardText}`;
            }
        }
        window.app.loadStylesToForm(initialInfo.styles);

        let renderTimeout, filenameTimeout;
        window.app.easyMDE.codemirror.on("change", () => {
            clearTimeout(renderTimeout);
            clearTimeout(filenameTimeout);
            renderTimeout = setTimeout(() => window.app.renderPreview(), 300);
            filenameTimeout = setTimeout(() => window.app.updateFilename(), 500);
        });

        document.getElementById('open-file-btn').addEventListener('click', () => window.pywebview.api.open_file_dialog());
        document.getElementById('template-presets-btn').addEventListener('click', () => { document.getElementById('template-dialog-overlay').classList.add('show'); });
        document.getElementById('clear-btn').addEventListener('click', () => { window.app.easyMDE.value(''); window.app.easyMDE.codemirror.focus(); });
        document.getElementById('paste-btn').addEventListener('click', async () => {
            const content = await window.pywebview.api.get_clipboard_content();
            const processedContent = await window.pywebview.api.process_text(content);
            window.app.easyMDE.value(processedContent);
            window.app.renderPreview();
            await window.app.updateFilename();
        });

        const copyHandler = (targetApp) => {
            const content = window.app.easyMDE.value();
            if (!content.trim()) {
                window.app.showNotification("内容为空，无法复制。", "info");
                return;
            }
            window.app.showNotification(`正在通过 ${targetApp.toUpperCase()} 后台处理，请稍候...`, "info");
            const currentStyles = window.app.readStylesFromForm();
            window.pywebview.api.copy_via_office_app(content, currentStyles, targetApp);
        };

        document.getElementById('copy-to-word-btn').addEventListener('click', () => copyHandler('word'));
        document.getElementById('copy-to-wps-btn').addEventListener('click', () => copyHandler('wps'));
        
        document.getElementById('select-folder-btn').addEventListener('click', async () => {
            const newPath = await window.pywebview.api.select_export_directory();
            if (newPath) {
                window.app.folderPathEl.textContent = newPath;
                window.app.folderPathEl.title = newPath;
            }
        });
        document.getElementById('save-file-btn').addEventListener('click', () => {
            const content = window.app.easyMDE.value();
            const directory = window.app.folderPathEl.textContent;
            const filename = window.app.fileNameInputEl.value;
            if (!content.trim() || !filename.trim()) {
                window.app.showNotification("内容和文件名不能为空。", "info");
                return;
            }
            const currentStyles = window.app.readStylesFromForm();
            window.pywebview.api.save_word_document(content, directory, filename, currentStyles);
        });

        function setupModal(overlayId, openBtnId) {
            const overlay = document.getElementById(overlayId);
            if (openBtnId) {
                document.getElementById(openBtnId).addEventListener('click', () => overlay.classList.add('show'));
            }
            overlay.addEventListener('click', (e) => {
                if (e.target === overlay || e.target.classList.contains('close-btn') || e.target.classList.contains('cancel-btn')) {
                    overlay.classList.remove('show');
                }
            });
        }
        setupModal('about-dialog-overlay', 'about-btn');
        setupModal('export-success-dialog-overlay');

        document.getElementById('open-file-btn-success').addEventListener('click', () => {
            window.pywebview.api.open_file(window.app.currentFilePath);
            document.getElementById('export-success-dialog-overlay').classList.remove('show');
        });
        document.getElementById('open-folder-btn-success').addEventListener('click', () => {
            window.pywebview.api.open_folder(window.app.currentFilePath);
            document.getElementById('export-success-dialog-overlay').classList.remove('show');
        });

        const styleOverlay = document.getElementById('style-settings-dialog-overlay');
        document.getElementById('style-settings-btn').addEventListener('click', () => {
            window.app.loadStylesToForm(JSON.parse(JSON.stringify(window.app.styles)));
            styleOverlay.classList.add('show');
        });
        document.getElementById('cancel-style-settings-btn').addEventListener('click', () => styleOverlay.classList.remove('show'));
        styleOverlay.addEventListener('click', (e) => {
            if (e.target === styleOverlay) styleOverlay.classList.remove('show');
        });
        document.getElementById('save-style-settings-btn').addEventListener('click', async () => {
            const newStyles = window.app.readStylesFromForm();
            await window.pywebview.api.save_styles(newStyles, 'custom');
            window.app.lastPreset = 'custom';
            window.app.templateStatus.textContent = '自定义样式';
            window.app.templateStatus.title = '自定义样式';
            window.app.showNotification("样式设置已保存。", "success");
            styleOverlay.classList.remove('show');
        });

        const templateOverlay = document.getElementById('template-dialog-overlay');
        let selectedPreset = window.app.lastPreset;
        document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
        document.querySelector(`.template-card[data-preset="${selectedPreset}"]`)?.classList.add('selected');
        templateOverlay.addEventListener('click', (e) => {
            if (e.target === templateOverlay) templateOverlay.classList.remove('show');
        });
        document.querySelectorAll('.template-card[data-preset]').forEach(card => {
            card.addEventListener('click', (e) => {
                const preset = e.currentTarget.dataset.preset;
                if (preset === 'custom') {
                    window.pywebview.api.select_template_dialog();
                    templateOverlay.classList.remove('show');
                } else {
                    document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                    e.currentTarget.classList.add('selected');
                    selectedPreset = preset;
                }
            });
        });
        document.getElementById('apply-template-btn').addEventListener('click', async () => {
            if (selectedPreset && selectedPreset !== 'custom') {
                const newStyles = presetStyles[selectedPreset];
                window.app.loadStylesToForm(newStyles);
                window.app.lastPreset = selectedPreset;
                await window.pywebview.api.save_styles(newStyles, selectedPreset);
                const cardText = document.querySelector(`.template-card[data-preset="${selectedPreset}"] span`).textContent.trim();
                window.app.showNotification(`已应用模板: ${cardText}`, 'success');
                window.app.templateStatus.textContent = cardText;
                window.app.templateStatus.title = `预设: ${cardText}`;
            }
            templateOverlay.classList.remove('show');
        });
        document.getElementById('cancel-template-btn').addEventListener('click', () => templateOverlay.classList.remove('show'));

        const tooltip = document.getElementById('template-tooltip');
        document.querySelectorAll('.template-card[data-preset]').forEach(card => {
            const presetName = card.dataset.preset;
            if (!presetName || presetName === 'custom') return;
            card.addEventListener('mouseenter', (e) => {
                const styles = presetStyles[presetName];
                if (!styles) return;
                let tooltipContent = '';
                const nameMapping = { body: '正文', h1: 'H1', h2: 'H2', h3: 'H3' };
                for (const el in styles) {
                    const s = styles[el];
                    tooltipContent += `<div class="tooltip-line"><span>${nameMapping[el]}: ${s.font}, ${s.size}pt</span><span class="tooltip-color-swatch" style="background-color: #${s.color};"></span></div>`;
                }
                tooltip.innerHTML = tooltipContent;
                tooltip.style.display = 'block';
            });
            card.addEventListener('mousemove', (e) => {
                tooltip.style.left = (e.clientX + 15) + 'px';
                tooltip.style.top = (e.clientY + 15) + 'px';
            });
            card.addEventListener('mouseleave', () => {
                tooltip.style.display = 'none';
            });
        });

        document.getElementById('text-processing-btn').addEventListener('click', () => window.app.showTextProcessingDialog());
        document.getElementById('save-text-processing-btn').addEventListener('click', () => window.app.saveTextProcessingSettings());
        document.getElementById('cancel-text-processing-btn').addEventListener('click', () => {
            document.getElementById('text-processing-dialog-overlay').classList.remove('show');
        });

        window.app.renderPreview();
        await window.app.updateFilename();

        document.getElementById('loader').style.opacity = '0';
        document.querySelector('.app-container').style.visibility = 'visible';
        document.querySelector('.app-container').style.opacity = '1';
        setTimeout(() => {
            document.getElementById('loader').style.display = 'none';
        }, 400);

        // 添加CSS样式
        const style = document.createElement('style');
        style.textContent = `
            .switch-label {
                display: flex;
                flex-direction: column;
                margin-bottom: 15px;
                cursor: pointer;
            }
            .checkbox-row {
                display: flex;
                align-items: center;
                margin-bottom: 4px;
            }
            .switch-text {
                font-weight: 500;
                margin-left: 8px;
            }
            .switch-description {
                font-size: 0.9em;
                color: var(--info-text);
                margin-bottom: 8px;
                margin-left: 24px;
            }
            .text-processing-options {
                margin: 20px 0;
            }
        `;
        document.head.appendChild(style);

    } catch (e) {
        alert('应用初始化失败，请检查日志。错误: ' + e.message);
        console.error(e);
    }
}
window.addEventListener('pywebviewready', initializeApp);
</script></body></html>
"""


def create_and_run_gui():
    """创建并运行WebView GUI"""
    logging.info("开始创建GUI...")

    # 获取主脚本所在目录
    main_script_file = getattr(sys.modules['__main__'], '__file__', '')
    if main_script_file:
        main_script_dir = os.path.dirname(os.path.abspath(main_script_file))
    else:
        main_script_dir = os.getcwd()  # 默认使用当前工作目录

    index_html_path = None
    try:
        wechat_qr_base64 = image_to_base64("wechat_qr.png")
        alipay_qr_base64 = image_to_base64("alipay_qr.png")

        # 检测系统主题
        is_dark = False
        try:
            is_dark = darkdetect.isDark()
            logging.info(f"检测到系统主题: {'暗色' if is_dark else '亮色'}")
        except Exception as e:
            logging.warning(f"无法检测系统主题: {e}")

        # 动态插入暗色主题相关的脚本
        theme_script = """
        <script>
            // 在文档加载前标记主题以避免闪烁
            document.documentElement.setAttribute('data-theme', '%(theme)s');
            document.documentElement.style.backgroundColor = '%(bg_color)s';
            document.documentElement.style.color = '%(text_color)s';
        </script>
        """ % {
            'theme': 'dark' if is_dark else 'light',
            'bg_color': '#252526' if is_dark else '#f5f5f5',
            'text_color': '#d4d4d4' if is_dark else '#242424'
        }
        
        final_html_content = html_template.replace(
            "__WECHAT_QR_PLACEHOLDER__", wechat_qr_base64
        ).replace(
            "__ALIPAY_QR_PLACEHOLDER__", alipay_qr_base64
        )
        
        # 在<head>标签末尾插入主题初始化脚本
        final_html_content = final_html_content.replace('</head>', f'{theme_script}</head>')

        index_html_path = os.path.join(main_script_dir, 'index.html')
        with open(index_html_path, 'w', encoding='utf-8') as f:
            f.write(final_html_content)
        logging.info(f"动态创建HTML文件于: {index_html_path}")

        api = Api()
        
        bg_color = '#252526' if is_dark else '#f5f5f5'

        # 先创建隐藏的窗口
        window = webview.create_window(
            'Markdown to Word Converter by Youkies',
            url=index_html_path,
            js_api=api,
            width=1400,
            height=900,
            resizable=True,
            min_size=(1200, 700),
            background_color=bg_color,
            hidden=True,
            x=None,  # 让窗口管理器自动定位窗口，防止位置异常
            y=None
        )

        backend_api_module.window = window

        def apply_dark_theme():
            """应用暗色主题到窗口标题栏"""
            if is_dark and platform.system() == "Windows":
                logging.info("正在应用暗色标题栏...")
                try:
                    # 通过FindWindowW查找窗口句柄
                    hwnd = None
                    try:
                        hwnd = ctypes.windll.user32.FindWindowW(None, window.title)
                        logging.info(f"通过FindWindowW获取句柄: {hwnd}")
                    except Exception:
                        logging.warning("通过FindWindowW获取窗口句柄失败")
            
                    # 如果已找到句柄，应用暗色主题
                    if hwnd:
                        set_dark_title_bar(hwnd)
                        logging.info(f"成功应用暗色标题栏到窗口 {hwnd}")
                    else:
                        logging.warning("无法找到窗口句柄，将继续但不应用暗色标题栏")
                except Exception as e:
                    # 即使出错也继续程序
                    logging.error(f"应用暗色标题栏时出错: {e}", exc_info=True)
                    logging.info("继续程序执行，忽略暗色标题栏应用错误")

        def on_loaded():
            """DOM加载完成后的处理"""
            logging.info("DOM 加载完成，准备显示窗口...")
            
            # 先尝试应用暗色主题，但不依赖于结果
            if is_dark:
                apply_dark_theme()
                
                # 通过JS再次确保使用暗色主题
                try:
                    window.evaluate_js("""
                        document.documentElement.setAttribute('data-theme', 'dark');
                        document.documentElement.style.backgroundColor = '#252526';
                        document.body.style.backgroundColor = '#252526';
                        document.body.style.color = '#d4d4d4';
                    """)
                except Exception as e:
                    logging.warning(f"应用JS暗色样式失败: {e}")
            
            # 不使用复杂的JS回调机制，直接显示窗口
            # 这样可以避免潜在的回调问题导致窗口不显示
            logging.info("正在显示窗口...")
            try:
                # 这里简化窗口显示逻辑，直接调用show()
                window.show()
                logging.info("窗口显示命令已执行")
                
                # 在窗口显示后，设置适当的加载状态
                try:
                    window.evaluate_js("""
                        setTimeout(function() {
                            document.getElementById('loader').style.opacity = '0';
                            document.querySelector('.app-container').style.visibility = 'visible';
                            document.querySelector('.app-container').style.opacity = '1';
                            setTimeout(() => {
                                document.getElementById('loader').style.display = 'none';
                            }, 400);
                        }, 100);
                    """)
                except Exception as e:
                    logging.warning(f"设置窗口显示状态失败: {e}")
            except Exception as e:
                logging.error(f"显示窗口失败: {e}", exc_info=True)

        def on_closed():
            """窗口关闭时的清理工作"""
            logging.info("窗口关闭，开始执行清理任务...")
            api.cleanup_on_exit()
            if index_html_path and os.path.exists(index_html_path):
                try:
                    os.unlink(index_html_path)
                    logging.info(f"已删除临时HTML文件: {index_html_path}")
                except Exception as e:
                    logging.error(f"删除临时HTML文件失败 {index_html_path}: {e}")

        # 注册事件
        window.events.loaded += on_loaded
        window.events.closed += on_closed

        logging.info("启动 Webview...")
        webview.start(debug=False)

    except Exception as e:
        logging.critical("GUI 创建过程中发生致命错误: ", exc_info=True)
        if index_html_path and os.path.exists(index_html_path):
            os.unlink(index_html_path)
        raise e