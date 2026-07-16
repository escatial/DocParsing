# Changelog

All notable changes to this project will be documented in this file.

## [1.0.2] - 2026-07

### Added
- Markdown 预览区支持 KaTeX 公式渲染（行内 `$...$` / 块级 `$$...$$`）
- 公式输出策略：保留 LaTeX 字符串原样，用户可手动复制到 MathType

### Dependencies
- 前端新增：`remark-math@6`、`rehype-katex@7`、`katex@0.16`

## [1.0.1] - 2026-07

### Added
- Word 脚注自动化：数字引用 `[N]` / 区间 `[N-M]` / 并列 `[N,M,K]` / APA `(作者等, 年份)`
- 自动探测引用风格并路由到对应处理管线
- 孤儿文献（参考文献列表中未在正文出现的）自动生成提示型脚注
- 文章标题多重兜底提取（含首段启发式、中英混合作者、纯装饰字符过滤）

### Fixed
- `[N]` 数字脚注在 ZIP md 名字非 `full.md` / 风格探测失败时被跳过
- Word 脚注引用数字在部分客户端不是上标（显式 `w:vertAlign="superscript"`）
- 正文解析卡在「上传文件」阶段（前/后端接口异步解耦）
- `Content-Disposition` 中文文件名导致下载响应头异常（ASCII fallback + UTF-8 `filename*`）
- 批量模式下载 Word 按钮不可见
- 标题提取对纯 `#` 装饰字符误识别
- ASCII 方括号 `[Title]` 没被清理为 `Title`

### Performance
- `158 KB Markdown` 完整解析：`12.6 ms`
- 单元测试 25+ / 端到端 / 边界 / 并发 测试稳定通过
- 生产构建产物：`116 KB gzip` (`vite build`)

### Docs
- README 更新核心功能 / Bug 修复记录

## [1.0.0] - 2026-07

### Added
- 初始发布：基于 MinerU 云 API 的高精度文档解析 Web 应用
- FastAPI 后端代理层 + React + TypeScript + Vite 前端
- 多格式输入（PDF/Word/PPT/Excel/图片/HTML）
- 5 段进度可视化 + 4 段子阶段进度
- 单文件/批量解析（最多 10 个）
- 任务取消（单文件/整批）
- Markdown 预览 + 大纲提取 + 参考文献识别
- 响应式布局（桌面双列 / 移动单列）
- ErrorBoundary 错误兜底