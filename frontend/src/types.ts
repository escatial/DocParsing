// 类型定义
export interface ParseResponse {
  internal_id: string;
  batch_id: string;
}

export type TaskState =
  | 'pending'
  | 'running'
  | 'done'
  | 'failed'
  | 'converting'
  | 'cancelled';

export type SubStage =
  | 'queued'
  | 'parsing_page'
  | 'inferencing'
  | 'assembling'
  | 'done'
  | 'failed'
  | 'cancelled'
  | 'unknown';

export type ExtraFormat = 'docx' | 'html' | 'latex';

export interface TaskInfo {
  task_id: string;
  state: TaskState;
  extracted_pages: number;
  total_pages: number;
  full_zip_url: string | null;
  err_msg: string;
  model_version: string;
  cancelled: boolean;
  sub_stage: SubStage | string;
}

// ============ 批量任务类型 ============
export interface BatchFileItem {
  internal_id: string;
  filename: string;
  size: number;
  state: TaskState;
  sub_stage: SubStage | string;
  extracted_pages: number;
  total_pages: number;
  full_zip_url: string | null;
  err_msg: string;
  cancelled: boolean;
}

export interface BatchInfo {
  internal_id: string;
  state: 'pending' | 'running' | 'done' | 'failed' | 'cancelled' | 'partial';
  file_count: number;
  done_count: number;
  failed_count: number;
  cancelled_count: number;
  cancelled_all: boolean;
  files: BatchFileItem[];
  created_at: number;
}

export interface BatchSubmitFile {
  internal_id: string;
  filename: string;
  size: number;
}

export interface BatchParseResponse {
  internal_id: string;
  batch_id: string;
  files: BatchSubmitFile[];
}

// ============ 解析结果 ============
export interface TocItem {
  level: number;
  text: string;
}

export interface ParseResult {
  filename: string;
  title: string;             // 文章标题（用于下载文件名）
  markdown: string;
  files: ResultFile[];
  available_formats: ExtraFormat[];
  toc: TocItem[];
  references: string[];
}

export interface ResultFile {
  name: string;
  format: ExtraFormat;
  size: number;
  content_base64: string;
}

// 工作流阶段（前端 UI 状态机）
export type WorkflowStage =
  | 'idle'
  | 'preview'
  | 'uploading'
  | 'parsing'
  | 'converting'
  | 'done'
  | 'failed'
  | 'cancelled';

// 解析模式
export type ParseMode = 'single' | 'batch';