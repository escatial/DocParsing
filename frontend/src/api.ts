import type {
  ParseResponse,
  TaskInfo,
  ParseResult,
  ExtraFormat,
  BatchParseResponse,
  BatchInfo,
} from './types';

const BASE = '/api';

export interface SubmitOptions {
  modelVersion: 'vlm' | 'pipeline';
  isOcr: boolean;
  enableFormula: boolean;
  enableTable: boolean;
}

// 单文件提交
export async function submitParse(
  file: File,
  opts: SubmitOptions
): Promise<ParseResponse> {
  const form = new FormData();
  form.append('file', file);
  form.append('model_version', opts.modelVersion);
  form.append('is_ocr', String(opts.isOcr));
  form.append('enable_formula', String(opts.enableFormula));
  form.append('enable_table', String(opts.enableTable));

  const res = await fetch(`${BASE}/parse`, { method: 'POST', body: form });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ msg: res.statusText }));
    throw new Error(err.msg || '提交失败');
  }
  return res.json();
}

// 批量提交（2-10 个文件）
export async function submitBatch(
  files: File[],
  opts: SubmitOptions
): Promise<BatchParseResponse> {
  const form = new FormData();
  files.forEach((f) => form.append('files', f));
  form.append('model_version', opts.modelVersion);
  form.append('is_ocr', String(opts.isOcr));
  form.append('enable_formula', String(opts.enableFormula));
  form.append('enable_table', String(opts.enableTable));

  const res = await fetch(`${BASE}/parse/batch`, { method: 'POST', body: form });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ msg: res.statusText }));
    throw new Error(err.msg || '批量提交失败');
  }
  return res.json();
}

export async function getTask(internalId: string): Promise<TaskInfo> {
  const res = await fetch(`${BASE}/task/${internalId}`);
  if (!res.ok) throw new Error('查询任务失败');
  return res.json();
}

export async function getBatch(batchInternalId: string): Promise<BatchInfo> {
  const res = await fetch(`${BASE}/batch/${batchInternalId}`);
  if (!res.ok) throw new Error('查询批量任务失败');
  return res.json();
}

export async function getResult(internalId: string): Promise<ParseResult> {
  const res = await fetch(`${BASE}/task/${internalId}/download`);
  if (!res.ok) throw new Error('获取结果失败');
  return res.json();
}

// 取消单文件任务
export async function cancelTask(internalId: string): Promise<void> {
  await fetch(`${BASE}/task/${internalId}/cancel`, { method: 'POST' });
}

// 取消整批任务
export async function cancelBatchTask(batchInternalId: string): Promise<void> {
  await fetch(`${BASE}/batch/${batchInternalId}/cancel`, { method: 'POST' });
}

export function downloadFormatUrl(internalId: string, format: ExtraFormat): string {
  return `${BASE}/task/${internalId}/format/${format}`;
}