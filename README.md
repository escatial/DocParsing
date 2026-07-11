# MinerU DocParser · 文档解析 Web 应用

> 基于 [MinerU](https://github.com/opendatalab/MinerU) 云端 API 构建的高精度文档解析 Web 应用。
> 前端：React + TypeScript + Vite + Tailwind + React Query + Lucide Icons
> 后端：FastAPI（作为 MinerU API 的安全代理）

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Python](https://img.shields.io/badge/python-3.11+-blue)
![Node](https://img.shields.io/badge/node-18+-green)

---

## ✨ 核心功能

### 解析能力
- **多格式输入**：PDF、Word、PPT、Excel、图片（png/jpg/webp 等）、HTML
- **多模型支持**：`vlm`（高精度 90+） / `pipeline`（CPU 友好）
- **可选能力**：OCR 识别、公式识别（LaTeX）、表格识别（HTML）

### 输出能力
- **Markdown 预览**：在线渲染完整 Markdown（含表格/代码块/列表/GFM 语法）
- **大纲提取**：自动从文档中提取章节结构（按 Markdown 标题层级）
- **参考文献识别**：自动识别"参考文献"章节并提取条目
- **Word 导出**：下载 .docx 文件，可在 Word/WPS 中打开编辑

### 工作流特性
- **单文件 / 批量解析**（最多 10 个文件）
- **确认按钮工作流**：选择文件 → 预览确认 → 提交解析（避免误操作）
- **5 段进度可视化**：预览 → 上传 → 解析 → 转换 → 完成
- **4 段子阶段进度**：MinerU 内部流程（排队 / 页面解析 / 模型推理 / 结果组装）
- **可取消任务**：单文件取消 + 整批取消
- **响应式布局**：桌面双列 / 移动端单列自适应
- **错误兜底**：ErrorBoundary 捕获异常并友好展示

## 📁 项目结构

```
文档解析/
├── backend/                # FastAPI 后端
│   ├── main.py             # 主程序（含详细注释）
│   ├── requirements.txt    # Python 依赖
│   └── .env.example        # 环境变量模板
└── frontend/               # React 前端
    ├── package.json
    ├── vite.config.ts      # 含 /api 代理配置
    ├── tailwind.config.js
    ├── tsconfig.json
    ├── index.html          # 含 favicon
    ├── public/
    │   └── favicon.svg     # 浏览器图标
    └── src/
        ├── main.tsx        # 入口
        ├── App.tsx         # 主组件（含 5 段进度条 + 双列布局）
        ├── ErrorBoundary.tsx
        ├── api.ts          # API 客户端
        ├── types.ts        # TypeScript 类型
        ├── queryClient.ts  # React Query 配置
        └── index.css       # Tailwind + Markdown 样式
```

## 🚀 启动步骤

### 1. 创建 conda 环境
```bash
conda create -n DocParsing python=3.11 -y
conda activate DocParsing
```

### 2. 安装并启动后端
```bash
cd backend
pip install -r requirements.txt
python -m uvicorn main:app --reload --port 8000
```

> Token 已内置在 `main.py` 中。如需自定义：
> ```bash
> # backend/.env
> MINERU_TOKEN=your_token_here
> ```

### 3. 安装并启动前端
```bash
cd frontend
npm install
npm run dev
```

浏览器访问 [http://localhost:5173](http://localhost:5173)

> 前端通过 Vite Proxy 将 `/api/*` 转发到 `http://127.0.0.1:8000`，开发期无需关心跨域。

## 🔄 API 流程（后端 → MinerU 云端）

```
1. POST  /api/v4/file-urls/batch    申请批量上传链接
2. PUT   <upload_url>               上传文件二进制（不要设 Content-Type）
3. GET   /api/v4/extract-results/batch/{batch_id}
                                    轮询批量任务状态
4. GET   <full_zip_url>             解析完成后下载 ZIP
5. 后端解压 ZIP，提取:
   - full.md → Markdown 预览
   - content_list.json → 大纲 + 参考文献
   - *.docx → Word 下载
```

后端核心代理层把 MinerU API 的 `batch_id` 和单文件 `task_id` 做了统一封装：
- 上传成功后 MinerU 自动创建任务
- 轮询必须用 `/extract-results/batch/{batch_id}` 端点
- 响应结构是 `extract_result[0]` 列表

## 🛠 技术栈

### 后端
| 技术 | 用途 |
|---|---|
| `FastAPI + uvicorn` | 异步 API 服务 |
| `httpx` | 异步 HTTP 客户端，调用 MinerU |
| `asyncio.gather` | 批量文件并行上传 |
| `pydantic v2` | 数据模型 + 验证 |
| `zipfile + io` | 流式解压 MinerU 结果 ZIP |
| `正则表达式` | 从 Markdown 兜底提取大纲和参考文献 |

### 前端
| 技术 | 用途 |
|---|---|
| `React 18 + TypeScript` | 主流技术栈 |
| `Vite` | 快速构建 + 内置 `/api` 代理 |
| `Tailwind CSS` | 原子化样式 |
| `React Query (TanStack Query)` | 轮询 + 缓存 + 错误处理 |
| `react-markdown + remark-gfm + rehype-sanitize` | Markdown 渲染（GFM + XSS 防护） |
| `lucide-react` | 现代图标库 |
| `useState + useEffect + useRef` | 工作流状态机 |
| `ErrorBoundary` | 异常兜底 |

## 🐛 Bug 修复记录

| Bug | 原因 | 修复 |
|---|---|---|
| `404 task not found` | 单文件 API 查询路径用了 `/extract/task/{id}`，但批量上传的 batch_id 必须用 `/extract-results/batch/{id}` | 切换查询端点 + 解析 `extract_result[0]` |
| `500 UnicodeEncodeError` | ZIP 下载时文件名含中文 | 使用 RFC 5987 `filename*=UTF-8''...` |
| `Pydantic UserWarning model_*` | 字段名 `model_version` 触发保护命名空间 | 添加 `model_config = {"protected_namespaces": ()}` |
| `react-markdown sentence undefined` | react-markdown v9 + remark-gfm v4 + react 18 偶发 | 锁定精确版本 + 加 `rehype-sanitize` |
| 持续 500 循环 | MinerU 临时不可用 | try-except 返回上次缓存状态 |
| 渲染时 setState 崩溃 | `if (X) setState(...)` 写在组件函数体内 | 改为 `useEffect` |
| `STATE_PILL[undefined].bg` 崩溃 | 后端返回未知 state | 添加 `getPill()` 兜底函数 |
| 批量只能预览第一个文件 | 文件区不可点击 | 改为可点击切换的 `<button>` |

## 📋 已知限制

- 任务状态保存在内存 `dict`，后端重启会丢失
- 仅展示 `full.md` + `docx`，未展示 `middle.json / model.json / content_list.json`
- 无用户登录/鉴权（适合内网演示）

## 🔮 后续可扩展

- 任务持久化（PostgreSQL + Prisma / SQLAlchemy）
- 解析历史列表 + 检索
- middle.json / model.json 单独 Tab 展示
- 公式（KaTeX）渲染
- 用户登录 + 配额管理
- WebSocket 推送替代轮询
- 接入本地 `mineru-api` 服务，去掉云端依赖

## 📜 License

MIT - 仅用于学习和演示