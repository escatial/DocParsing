import { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import { useQuery } from '@tanstack/react-query';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import rehypeSanitize from 'rehype-sanitize';
import {
  Upload,
  FileText,
  Loader2,
  CheckCircle2,
  XCircle,
  Download,
  AlertCircle,
  FileType,
  Sparkles,
  Play,
  Cpu,
  FileCheck,
  Ban,
  ListTree,
  BookMarked,
  Plus,
  X,
  ListChecks,
  ScanSearch,
  Brain,
  Package,
} from 'lucide-react';
import {
  submitParse,
  submitBatch,
  getTask,
  getBatch,
  getResult,
  cancelTask,
  cancelBatchTask,
  downloadFormatUrl,
} from './api';
import type {
  TaskInfo,
  ParseResult,
  WorkflowStage,
  ExtraFormat,
  SubStage,
  ParseMode,
  BatchInfo,
  BatchFileItem,
  TocItem,
} from './types';

// 主阶段配置
const STAGE_LIST: { id: WorkflowStage; label: string; icon: typeof Upload }[] = [
  { id: 'preview', label: '文件预览', icon: FileText },
  { id: 'uploading', label: '上传文件', icon: Upload },
  { id: 'parsing', label: 'MinerU 解析', icon: Cpu },
  { id: 'converting', label: '格式转换', icon: Sparkles },
  { id: 'done', label: '完成', icon: FileCheck },
];

// 子阶段
const SUB_STAGE_LIST: {
  id: SubStage;
  label: string;
  desc: string;
  icon: typeof Cpu;
}[] = [
  { id: 'queued', label: '排队等待', desc: 'MinerU 调度资源中', icon: Cpu },
  { id: 'parsing_page', label: '页面解析', desc: 'OCR / 版面分析 / 切分', icon: ScanSearch },
  { id: 'inferencing', label: '模型推理', desc: 'VLM/Pipeline 模型分析', icon: Brain },
  { id: 'assembling', label: '结果组装', desc: '拼接 Markdown / 转换格式', icon: Package },
];

const MAX_BATCH_FILES = 10;

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
}

const STATE_PILL: Record<string, { bg: string; text: string; label: string }> = {
  pending: { bg: 'bg-slate-100', text: 'text-slate-600', label: '排队中' },
  running: { bg: 'bg-blue-100', text: 'text-blue-700', label: '解析中' },
  converting: { bg: 'bg-purple-100', text: 'text-purple-700', label: '转换中' },
  done: { bg: 'bg-emerald-100', text: 'text-emerald-700', label: '完成' },
  failed: { bg: 'bg-red-100', text: 'text-red-700', label: '失败' },
  cancelled: { bg: 'bg-orange-100', text: 'text-orange-700', label: '已取消' },
};

const getPill = (state: string | undefined | null) =>
  STATE_PILL[state ?? ''] ?? STATE_PILL.pending;

// 把 Markdown 内容按一级/二级标题切成段落（用于大纲展示）
function buildTocFromMarkdown(md: string): TocItem[] {
  const toc: TocItem[] = [];
  const lines = md.split('\n');
  for (const line of lines) {
    const m = line.match(/^(#{1,3})\s+(.+?)\s*$/);
    if (m && m[2] && !m[2].startsWith('```')) {
      toc.push({ level: m[1].length, text: m[2] });
    }
  }
  return toc;
}

export default function App() {
  // ===== 工作流状态 =====
  const [stage, setStage] = useState<WorkflowStage>('idle');
  const [mode, setMode] = useState<ParseMode>('single');
  const [pendingFiles, setPendingFiles] = useState<File[]>([]);
  const [internalId, setInternalId] = useState<string | null>(null);
  const [batchInternalId, setBatchInternalId] = useState<string | null>(null);
  const [selectedFileId, setSelectedFileId] = useState<string | null>(null);
  const [finalResults, setFinalResults] = useState<Record<string, ParseResult>>({});
  const [submitError, setSubmitError] = useState<string | null>(null);
  const [submitting, setSubmitting] = useState(false);
  const [cancelling, setCancelling] = useState(false);

  // 选项
  const [modelVersion, setModelVersion] = useState<'vlm' | 'pipeline'>('vlm');
  const [isOcr, setIsOcr] = useState(false);
  const [enableFormula, setEnableFormula] = useState(true);
  const [enableTable, setEnableTable] = useState(true);

  const [dragging, setDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ===== 单文件轮询 =====
  const { data: taskInfo, error: taskError } = useQuery<TaskInfo>({
    queryKey: ['task', internalId],
    queryFn: () => getTask(internalId!),
    enabled: mode === 'single' && !!internalId && stage !== 'done' && stage !== 'cancelled',
    refetchInterval: 2000,
    retry: false,
  });

  // ===== 批量轮询 =====
  const { data: batchInfo, error: batchError } = useQuery<BatchInfo>({
    queryKey: ['batch', batchInternalId],
    queryFn: () => getBatch(batchInternalId!),
    enabled: mode === 'batch' && !!batchInternalId && stage !== 'done' && stage !== 'cancelled',
    refetchInterval: 2000,
    retry: false,
  });

  // 单文件状态机
  useEffect(() => {
    if (mode !== 'single' || !taskInfo) return;
    if (stage === 'done' || stage === 'cancelled' || stage === 'failed') return;
    if (taskInfo.state === 'failed') setStage('failed');
    else if (taskInfo.state === 'cancelled') setStage('cancelled');
    else if (taskInfo.state === 'done') {
      if (stage === 'converting') setStage('done');
      else setStage('converting');
    } else if (taskInfo.state === 'converting' && stage === 'parsing') setStage('converting');
    else if (taskInfo.state === 'running' && stage === 'uploading') setStage('parsing');
  }, [taskInfo?.state, taskInfo?.cancelled, mode, stage]);

  // 批量状态机
  useEffect(() => {
    if (mode !== 'batch' || !batchInfo) return;
    if (stage === 'done' || stage === 'cancelled') return;
    if (batchInfo.state === 'cancelled') setStage('cancelled');
    else if (batchInfo.state === 'done') setStage('done');
    else if (batchInfo.state === 'partial') {
      const stillRunning = batchInfo.files.some(
        (f) => !['done', 'failed', 'cancelled'].includes(f.state)
      );
      if (!stillRunning) setStage('done');
    }
  }, [batchInfo?.state, batchInfo?.files, mode, stage]);

  // 批量进入 done 时：自动选中第一个 done 的文件（防止 selectedFileId 为 null）
  useEffect(() => {
    if (mode !== 'batch' || stage !== 'done' || !batchInfo) return;
    if (selectedFileId) return;
    const firstDone = batchInfo.files.find((f) => f.state === 'done');
    if (firstDone) setSelectedFileId(firstDone.internal_id);
  }, [mode, stage, batchInfo, selectedFileId]);

  // 单文件结果
  const { data: singleResult } = useQuery<ParseResult>({
    queryKey: ['result', internalId],
    queryFn: () => getResult(internalId!),
    enabled: mode === 'single' && stage === 'done' && !!internalId,
    refetchOnMount: 'always',
    retry: false,
  });

  // 把单文件结果写入 finalResults（统一数据源）
  useEffect(() => {
    if (singleResult && internalId && mode === 'single' && !finalResults[internalId]) {
      setFinalResults((prev) => ({ ...prev, [internalId]: singleResult }));
    }
  }, [singleResult, internalId, mode, finalResults]);

  // 批量：拉取选中文件的结果
  const { data: batchSelectedResult } = useQuery<ParseResult>({
    queryKey: ['result', selectedFileId],
    queryFn: () => getResult(selectedFileId!),
    enabled:
      mode === 'batch' &&
      stage === 'done' &&
      !!selectedFileId &&
      !finalResults[selectedFileId],
    refetchOnMount: 'always',
    retry: false,
  });

  // 把批量选中的结果存到 finalResults
  useEffect(() => {
    if (batchSelectedResult && selectedFileId && !finalResults[selectedFileId]) {
      setFinalResults((prev) => ({ ...prev, [selectedFileId]: batchSelectedResult }));
    }
  }, [batchSelectedResult, selectedFileId, finalResults]);

  // ===== 文件选择 =====
  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    const files = Array.from(e.dataTransfer.files);
    if (files.length > 0) selectFiles(files);
  }, []);

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files ? Array.from(e.target.files) : [];
    if (files.length > 0) selectFiles(files);
  };

  const selectFiles = (files: File[]) => {
    const allowed = [
      'pdf', 'png', 'jpg', 'jpeg', 'jp2', 'webp', 'gif', 'bmp',
      'doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx', 'html',
    ];
    const valid: File[] = [];
    for (const f of files) {
      if (f.size > 200 * 1024 * 1024) {
        setSubmitError(`文件 ${f.name} 超过 200MB 限制`);
        continue;
      }
      const ext = (f.name.split('.').pop() ?? '').toLowerCase();
      if (!allowed.includes(ext)) {
        setSubmitError(`不支持的文件类型: .${ext}`);
        continue;
      }
      valid.push(f);
    }
    if (valid.length === 0) return;
    if (valid.length > MAX_BATCH_FILES) {
      setSubmitError(`批量最多 ${MAX_BATCH_FILES} 个文件，当前 ${valid.length} 个`);
      return;
    }
    setSubmitError(null);
    setMode(valid.length === 1 ? 'single' : 'batch');
    setPendingFiles(valid);
    setStage('preview');
  };

  const addMoreFiles = () => fileInputRef.current?.click();
  const removeFile = (idx: number) => {
    setPendingFiles((prev) => {
      const next = prev.filter((_, i) => i !== idx);
      if (next.length === 0) { setStage('idle'); setMode('single'); }
      else if (next.length === 1) setMode('single');
      return next;
    });
  };

  const confirmParse = async () => {
    if (pendingFiles.length === 0) return;
    setSubmitting(true);
    setSubmitError(null);
    setFinalResults({});
    setSelectedFileId(null);
    setStage('uploading');
    try {
      const opts = { modelVersion, isOcr, enableFormula, enableTable };
      if (pendingFiles.length === 1) {
        const res = await submitParse(pendingFiles[0], opts);
        setInternalId(res.internal_id);
        setBatchInternalId(null);
        setStage('parsing');
      } else {
        const res = await submitBatch(pendingFiles, opts);
        setBatchInternalId(res.internal_id);
        setInternalId(null);
        setStage('parsing');
        if (res.files.length > 0) setSelectedFileId(res.files[0].internal_id);
      }
    } catch (e) {
      setSubmitError(e instanceof Error ? e.message : String(e));
      setStage('preview');
    } finally {
      setSubmitting(false);
    }
  };

  const cancelCurrentTask = async () => {
    if (!internalId || cancelling) return;
    if (!confirm('确定要取消此任务吗？')) return;
    setCancelling(true);
    try {
      await cancelTask(internalId);
      setStage('cancelled');
    } catch (e) {
      alert('取消失败: ' + (e instanceof Error ? e.message : String(e)));
    } finally {
      setCancelling(false);
    }
  };

  const cancelCurrentBatch = async () => {
    if (!batchInternalId || cancelling) return;
    if (!confirm('确定要取消整批任务吗？')) return;
    setCancelling(true);
    try {
      await cancelBatchTask(batchInternalId);
      setStage('cancelled');
    } catch (e) {
      alert('取消失败: ' + (e instanceof Error ? e.message : String(e)));
    } finally {
      setCancelling(false);
    }
  };

  const cancelOneInBatch = async (fileInternalId: string) => {
    if (!confirm('确定取消此文件吗？')) return;
    try { await cancelTask(fileInternalId); }
    catch (e) { alert('取消失败: ' + (e instanceof Error ? e.message : String(e))); }
  };

  const reset = () => {
    setStage('idle');
    setMode('single');
    setPendingFiles([]);
    setInternalId(null);
    setBatchInternalId(null);
    setSelectedFileId(null);
    setFinalResults({});
    setSubmitError(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  // ===== 派生 =====
  const currentStageIndex = useMemo(() => {
    if (stage === 'idle') return -1;
    if (stage === 'failed' || stage === 'cancelled')
      return STAGE_LIST.findIndex((s) => s.id === 'parsing');
    return STAGE_LIST.findIndex((s) => s.id === stage);
  }, [stage]);

  const isFailed = stage === 'failed' || !!taskError || !!batchError;
  const isCancelled = stage === 'cancelled';
  const isBusy = stage === 'uploading' || stage === 'parsing' || stage === 'converting';

  // 当前结果（单文件模式）
  const currentResult: ParseResult | null = mode === 'single' ? (singleResult ?? null) : null;

  // 批量选中文件的结果
  const selectedBatchResult: ParseResult | null = useMemo(() => {
    if (mode !== 'batch' || !selectedFileId) return null;
    return finalResults[selectedFileId] ?? null;
  }, [mode, selectedFileId, finalResults]);

  // 大纲与参考文献：优先使用后端返回，否则从 Markdown 提取
  const toc: TocItem[] = useMemo(() => {
    const r = currentResult ?? selectedBatchResult;
    if (!r) return [];
    if (r.toc && r.toc.length > 0) return r.toc;
    return buildTocFromMarkdown(r.markdown);
  }, [currentResult, selectedBatchResult]);

  const references: string[] = useMemo(() => {
    const r = currentResult ?? selectedBatchResult;
    return r?.references ?? [];
  }, [currentResult, selectedBatchResult]);

  const hasDocx = (currentResult?.available_formats ?? []).includes('docx');

  // 单文件进度
  const singleProgressPct =
    taskInfo && taskInfo.total_pages > 0
      ? Math.round((taskInfo.extracted_pages / taskInfo.total_pages) * 100)
      : null;

  // 批量选中文件
  const selectedBatchFile: BatchFileItem | null = useMemo(() => {
    if (mode !== 'batch' || !selectedFileId || !batchInfo) return null;
    return batchInfo.files.find((f) => f.internal_id === selectedFileId) ?? null;
  }, [mode, selectedFileId, batchInfo]);

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50">
      <div className="max-w-7xl mx-auto p-4 md:p-6">
        {/* 顶部标题 */}
        <header className="mb-4">
          <h1 className="text-2xl md:text-3xl font-bold text-slate-900 flex items-center gap-2">
            <FileText className="text-blue-600" size={28} />
            MinerU DocParser
            <span className="text-xs font-normal text-slate-400 bg-slate-100 px-2 py-0.5 rounded ml-1">v1.0.0</span>
          </h1>
          <p className="text-slate-500 mt-1 text-xs md:text-sm">
            高精度文档解析 · 基于 MinerU 云端 API · 支持单文件 / 批量（最多 {MAX_BATCH_FILES} 个）
          </p>
        </header>

        {/* ==================== 第一行：横向排布「文件区 + 选项区」 ==================== */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-3 mb-3">
          {/* 文件区 */}
          <section className="lg:col-span-7 bg-white rounded-xl p-4 shadow-sm border border-slate-200">
            <div className="flex items-center justify-between mb-3">
              <h2 className="font-semibold text-slate-800 flex items-center gap-2 text-sm md:text-base">
                <span className="w-5 h-5 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center text-xs">1</span>
                {stage === 'idle' || stage === 'preview' ? '选择文件' : '已选文件'}
                {mode === 'batch' && (
                  <span className="text-xs text-blue-600 bg-blue-50 px-2 py-0.5 rounded">
                    批量 · {pendingFiles.length} 个
                  </span>
                )}
              </h2>
              {stage === 'preview' && pendingFiles.length < MAX_BATCH_FILES && (
                <button onClick={addMoreFiles} className="text-xs text-blue-600 hover:underline flex items-center gap-1">
                  <Plus size={12} /> 添加更多
                </button>
              )}
            </div>

            {stage === 'idle' && (
              <>
                <div
                  className={`border-2 border-dashed rounded-lg p-4 md:p-5 text-center cursor-pointer transition-all ${
                    dragging ? 'border-blue-500 bg-blue-50 scale-[1.02]' : 'border-slate-300 hover:border-blue-400 hover:bg-slate-50'
                  }`}
                  onClick={() => fileInputRef.current?.click()}
                  onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
                  onDragLeave={() => setDragging(false)}
                  onDrop={onDrop}
                >
                  <Upload className={`mx-auto mb-2 ${dragging ? 'text-blue-500' : 'text-slate-400'}`} size={28} />
                  <p className="text-slate-700 text-sm font-medium">点击或拖拽文件到此处</p>
                  <p className="text-xs text-slate-400 mt-1">
                    1-{MAX_BATCH_FILES} 个 · 单个 ≤ 200MB · PDF / 图片 / Office / HTML
                  </p>
                </div>
                {submitError && (
                  <div className="mt-2 text-sm text-red-600 flex items-start gap-1.5 bg-red-50 p-2 rounded">
                    <AlertCircle size={14} className="mt-0.5 shrink-0" />
                    <span>{submitError}</span>
                  </div>
                )}
              </>
            )}

            {stage === 'preview' && (
              <div className="space-y-2">
                <div className="max-h-40 overflow-y-auto space-y-1.5">
                  {pendingFiles.map((f, idx) => (
                    <div key={idx} className="border border-slate-200 rounded-lg p-2 bg-slate-50 flex items-center gap-2">
                      <FileText size={14} className="text-blue-500 shrink-0" />
                      <div className="flex-1 min-w-0">
                        <p className="text-xs font-medium text-slate-800 truncate">{f.name}</p>
                        <p className="text-xs text-slate-400">{formatSize(f.size)}</p>
                      </div>
                      <button onClick={() => removeFile(idx)} className="text-slate-400 hover:text-red-500">
                        <X size={14} />
                      </button>
                    </div>
                  ))}
                </div>
                <div className="flex gap-2 pt-1">
                  <button
                    onClick={confirmParse}
                    disabled={submitting}
                    className="flex-1 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium py-1.5 px-3 rounded-lg flex items-center justify-center gap-1.5 disabled:opacity-50"
                  >
                    {submitting ? <><Loader2 size={14} className="animate-spin" />提交中…</>
                      : <><Play size={14} />确认解析{pendingFiles.length > 1 && ` (${pendingFiles.length})`}</>}
                  </button>
                  <button
                    onClick={() => { setPendingFiles([]); setStage('idle'); setMode('single'); }}
                    disabled={submitting}
                    className="text-sm text-slate-600 hover:text-slate-800 border border-slate-300 rounded-lg px-3 hover:bg-slate-100"
                  >
                    重选
                  </button>
                </div>
              </div>
            )}

            {(isBusy || isFailed || isCancelled || stage === 'done') && (
              <div className="space-y-1.5">
                {(batchInfo?.files ?? pendingFiles.map((f) => ({ filename: f.name, size: f.size, internal_id: '', state: taskInfo?.state ?? 'pending' }))).map((f: any, idx) => {
                  const fileState = batchInfo?.files?.[idx]?.state ?? taskInfo?.state;
                  const pill = getPill(fileState);
                  const isSelected = mode === 'batch' && selectedFileId === f.internal_id;
                  const canSwitch = mode === 'batch' && (fileState === 'done' || isSelected);
                  return (
                    <button
                      key={idx}
                      onClick={() => canSwitch && setSelectedFileId(f.internal_id)}
                      disabled={!canSwitch}
                      className={`w-full text-left border rounded-lg p-2 flex items-center gap-2 transition ${
                        isSelected
                          ? 'border-blue-400 bg-blue-50 ring-1 ring-blue-300'
                          : canSwitch
                          ? 'border-slate-200 bg-white hover:border-slate-300 hover:bg-slate-50 cursor-pointer'
                          : 'border-slate-200 bg-white opacity-90 cursor-default'
                      }`}
                    >
                      <FileText size={14} className={`shrink-0 ${isSelected ? 'text-blue-500' : 'text-slate-400'}`} />
                      <div className="flex-1 min-w-0">
                        <p className={`text-xs font-medium truncate ${isSelected ? 'text-blue-700' : 'text-slate-800'}`}>
                          {f.filename}
                        </p>
                        <p className="text-xs text-slate-400">{formatSize(f.size)}</p>
                      </div>
                      <span className={`text-xs px-2 py-0.5 rounded ${pill.bg} ${pill.text}`}>
                        {pill.label}
                      </span>
                    </button>
                  );
                })}
              </div>
            )}

            <input
              ref={fileInputRef}
              type="file"
              className="hidden"
              multiple
              accept=".pdf,.png,.jpg,.jpeg,.webp,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.html"
              onChange={handleFileInput}
            />
          </section>

          {/* 选项区 */}
          <section className={`lg:col-span-5 bg-white rounded-xl p-4 shadow-sm border border-slate-200 ${stage === 'preview' || stage === 'idle' ? '' : 'opacity-60'}`}>
            <h2 className="font-semibold text-slate-800 mb-3 flex items-center gap-2 text-sm md:text-base">
              <span className="w-5 h-5 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center text-xs">2</span>
              解析选项
            </h2>
            <div className="grid grid-cols-2 gap-3 text-sm">
              <div>
                <label className="block text-slate-600 mb-1 text-xs">模型版本</label>
                <select
                  value={modelVersion}
                  onChange={(e) => setModelVersion(e.target.value as 'vlm' | 'pipeline')}
                  className="w-full border border-slate-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-blue-200"
                  disabled={isBusy}
                >
                  <option value="vlm">vlm（推荐）</option>
                  <option value="pipeline">pipeline（CPU）</option>
                </select>
              </div>
              <div className="space-y-1.5 pt-5">
                <label className="flex items-center gap-1.5 text-xs cursor-pointer">
                  <input type="checkbox" checked={isOcr} onChange={(e) => setIsOcr(e.target.checked)} disabled={isBusy} className="rounded text-blue-600" />
                  <span>OCR</span>
                </label>
                <label className="flex items-center gap-1.5 text-xs cursor-pointer">
                  <input type="checkbox" checked={enableFormula} onChange={(e) => setEnableFormula(e.target.checked)} disabled={isBusy} className="rounded text-blue-600" />
                  <span>公式识别</span>
                </label>
                <label className="flex items-center gap-1.5 text-xs cursor-pointer">
                  <input type="checkbox" checked={enableTable} onChange={(e) => setEnableTable(e.target.checked)} disabled={isBusy} className="rounded text-blue-600" />
                  <span>表格识别</span>
                </label>
              </div>
            </div>

            {(stage === 'done' || stage === 'cancelled' || stage === 'failed') && (
              <button onClick={reset} className="mt-3 w-full bg-white hover:bg-slate-50 text-slate-700 text-sm font-medium py-1.5 px-3 rounded-lg border border-slate-300">
                解析新文件
              </button>
            )}
          </section>
        </div>

        {/* ==================== 进度条（仅在 busy 或 done 时显示） ==================== */}
        {(isBusy || stage === 'done' || isFailed || isCancelled) && (
          <section className="bg-white rounded-xl p-3 md:p-4 shadow-sm border border-slate-200 mb-3">
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-2">
                {STAGE_LIST.map((s, idx) => {
                  const Icon = s.icon;
                  const isCurrent = stage === s.id;
                  const isDone =
                    currentStageIndex > idx ||
                    (stage === 'done' && idx <= STAGE_LIST.findIndex((x) => x.id === 'done'));
                  const isCancelStage = stage === 'cancelled' && idx === STAGE_LIST.findIndex((x) => x.id === 'parsing');
                  const isFailedStage = stage === 'failed' && idx === STAGE_LIST.findIndex((x) => x.id === 'parsing');
                  return (
                    <div key={s.id} className="flex items-center">
                      <div className="flex flex-col items-center">
                        <div
                          className={`w-7 h-7 rounded-full flex items-center justify-center border-2 transition-all ${
                            isCancelStage ? 'bg-orange-100 border-orange-400 text-orange-600'
                            : isFailedStage ? 'bg-red-100 border-red-400 text-red-600'
                            : isCurrent ? 'bg-blue-100 border-blue-500 text-blue-600 ring-4 ring-blue-100 animate-pulse'
                            : isDone ? 'bg-emerald-100 border-emerald-500 text-emerald-600'
                            : 'bg-slate-50 border-slate-200 text-slate-400'
                          }`}
                        >
                          {isCancelStage ? <Ban size={14} /> : isFailedStage ? <XCircle size={14} /> : isDone ? <CheckCircle2 size={14} /> : <Icon size={12} />}
                        </div>
                        <span className={`text-[10px] mt-1 font-medium whitespace-nowrap ${
                          isCancelStage ? 'text-orange-600' : isCurrent ? 'text-blue-600' : isDone ? 'text-emerald-600' : 'text-slate-400'
                        }`}>
                          {s.label}
                        </span>
                      </div>
                      {idx < STAGE_LIST.length - 1 && (
                        <div className={`w-6 md:w-12 h-0.5 mx-1 transition-all ${currentStageIndex > idx ? 'bg-emerald-400' : 'bg-slate-200'}`} />
                      )}
                    </div>
                  );
                })}
              </div>

              {isBusy && (
                <button
                  onClick={mode === 'batch' ? cancelCurrentBatch : cancelCurrentTask}
                  disabled={cancelling}
                  className="text-xs text-orange-600 hover:text-orange-700 border border-orange-300 hover:bg-orange-50 rounded px-2 py-1 flex items-center gap-1 disabled:opacity-50"
                >
                  {cancelling ? <Loader2 size={12} className="animate-spin" /> : <Ban size={12} />}
                  {mode === 'batch' ? '取消整批' : '取消'}
                </button>
              )}
            </div>

            {/* 子阶段 + 详情 */}
            {isBusy && (
              <div className="flex flex-col md:flex-row gap-2 mt-2">
                {/* 批量：总体进度 */}
                {mode === 'batch' && batchInfo && (
                  <div className="flex-1 p-2 bg-slate-50 rounded">
                    <div className="flex items-center justify-between text-xs mb-1.5">
                      <span className="font-medium text-slate-700">
                        完成 {batchInfo.done_count} / {batchInfo.file_count}
                      </span>
                    </div>
                    <div className="h-1.5 bg-slate-200 rounded overflow-hidden flex">
                      <div className="h-full bg-emerald-500" style={{ width: `${(batchInfo.done_count / batchInfo.file_count) * 100}%` }} />
                      <div className="h-full bg-red-400" style={{ width: `${(batchInfo.failed_count / batchInfo.file_count) * 100}%` }} />
                      <div className="h-full bg-orange-400" style={{ width: `${(batchInfo.cancelled_count / batchInfo.file_count) * 100}%` }} />
                    </div>
                  </div>
                )}

                {/* 单文件进度条 */}
                {mode === 'single' && stage === 'parsing' && singleProgressPct !== null && (
                  <div className="flex-1 p-2 bg-slate-50 rounded">
                    <div className="flex items-center justify-between text-xs mb-1.5">
                      <span className="font-medium text-slate-700">
                        {taskInfo?.extracted_pages ?? 0} / {taskInfo?.total_pages ?? '?'} 页
                      </span>
                      <span className="text-slate-500">{singleProgressPct}%</span>
                    </div>
                    <div className="h-1.5 bg-slate-200 rounded overflow-hidden">
                      <div className="h-full bg-gradient-to-r from-blue-500 to-cyan-500 transition-all" style={{ width: `${singleProgressPct}%` }} />
                    </div>
                  </div>
                )}

                {/* 子阶段 */}
                {mode === 'single' && taskInfo && (
                  <CompactSubStage
                    state={
                      taskInfo.state === 'done' ? 'done'
                      : taskInfo.state === 'failed' ? 'failed'
                      : taskInfo.state === 'cancelled' ? 'cancelled'
                      : (taskInfo.sub_stage as SubStage) || 'queued'
                    }
                    extracted={taskInfo.extracted_pages}
                    total={taskInfo.total_pages}
                  />
                )}
                {mode === 'batch' && selectedBatchFile && (
                  <CompactSubStage
                    state={
                      selectedBatchFile.state === 'done' ? 'done'
                      : selectedBatchFile.state === 'failed' ? 'failed'
                      : selectedBatchFile.state === 'cancelled' ? 'cancelled'
                      : (selectedBatchFile.sub_stage as SubStage) || 'queued'
                    }
                    extracted={selectedBatchFile.extracted_pages}
                    total={selectedBatchFile.total_pages}
                  />
                )}
              </div>
            )}

            {/* 完成终态保留 */}
            {stage === 'done' && (
              <div className="text-xs text-emerald-600 flex items-center gap-1.5 mt-1">
                <CheckCircle2 size={14} />
                <span>
                  解析完成
                  {currentResult?.markdown && ` · 共 ${currentResult.markdown.length} 字符`}
                  {mode === 'batch' && batchInfo && ` · ${batchInfo.done_count}/${batchInfo.file_count} 个文件完成`}
                </span>
              </div>
            )}
            {isFailed && (
              <div className="text-xs text-red-600 flex items-center gap-1.5 mt-1">
                <XCircle size={14} />
                <span>{taskInfo?.err_msg || (taskError as Error)?.message || (batchError as Error)?.message || '解析失败'}</span>
              </div>
            )}
            {isCancelled && (
              <div className="text-xs text-orange-600 flex items-center gap-1.5 mt-1">
                <Ban size={14} />
                <span>任务已取消（MinerU 仍在云端执行）</span>
              </div>
            )}
          </section>
        )}

        {/* ==================== 第二行：双列布局（Markdown | 大纲+参考文献） ==================== */}
        {stage === 'done' && (currentResult || selectedBatchResult) && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-3">
            {/* 左：Markdown 预览 */}
            <section className="lg:col-span-8 bg-white rounded-xl p-4 md:p-5 shadow-sm border border-slate-200">
              <div className="flex items-center justify-between mb-3">
                <h2 className="font-semibold text-slate-800 flex items-center gap-2 text-sm md:text-base">
                  <FileText size={16} className="text-blue-600" />
                  Markdown 预览
                  <span className="text-xs text-slate-400 font-normal">
                    {((currentResult ?? selectedBatchResult)?.title) ||
                      ((currentResult ?? selectedBatchResult)?.filename) ||
                      ''}
                  </span>
                </h2>
                <div className="flex items-center gap-2">
                  {hasDocx && (currentResult || selectedBatchResult) && (
                    <a
                      href={downloadFormatUrl(
                        currentResult ? internalId! : selectedFileId!,
                        'docx'
                      )}
                      download
                      className="text-xs flex items-center gap-1 text-white bg-blue-600 hover:bg-blue-700 rounded px-2.5 py-1"
                    >
                      <FileType size={12} />
                      下载 Word
                    </a>
                  )}
                </div>
              </div>
              <article className="markdown-body max-w-none max-h-[70vh] overflow-y-auto">
                {(currentResult ?? selectedBatchResult)?.markdown ? (
                  <ReactMarkdown remarkPlugins={[remarkGfm]} rehypePlugins={[rehypeSanitize]}>
                    {(currentResult ?? selectedBatchResult)!.markdown}
                  </ReactMarkdown>
                ) : (
                  <p className="text-slate-400 italic">（解析结果为空）</p>
                )}
              </article>
            </section>

            {/* 右：大纲 + 参考文献 */}
            <aside className="lg:col-span-4 space-y-3">
              {/* 大纲 */}
              <section className="bg-white rounded-xl p-4 shadow-sm border border-slate-200">
                <h2 className="font-semibold text-slate-800 flex items-center gap-2 mb-3 text-sm">
                  <ListTree size={16} className="text-blue-600" />
                  大纲
                  <span className="text-xs text-slate-400 font-normal ml-auto">
                    {toc.length} 项
                  </span>
                </h2>
                {toc.length > 0 ? (
                  <ul className="space-y-1 max-h-72 overflow-y-auto text-sm">
                    {toc.map((item, idx) => (
                      <li
                        key={idx}
                        style={{ paddingLeft: `${(item.level - 1) * 12}px` }}
                        className={`text-slate-700 hover:text-blue-600 cursor-default truncate ${
                          item.level === 1 ? 'font-semibold text-slate-900' : ''
                        }`}
                        title={item.text}
                      >
                        <span className="text-slate-400 mr-1 text-xs">
                          {'└'.repeat(item.level - 1)}
                          {item.level === 1 ? '' : ''}
                        </span>
                        {item.text}
                      </li>
                    ))}
                  </ul>
                ) : (
                  <p className="text-xs text-slate-400 italic">未识别到章节标题</p>
                )}
              </section>

              {/* 参考文献 */}
              <section className="bg-white rounded-xl p-4 shadow-sm border border-slate-200">
                <h2 className="font-semibold text-slate-800 flex items-center gap-2 mb-3 text-sm">
                  <BookMarked size={16} className="text-blue-600" />
                  参考文献
                  <span className="text-xs text-slate-400 font-normal ml-auto">
                    {references.length} 条
                  </span>
                </h2>
                {references.length > 0 ? (
                  <ol className="space-y-2 max-h-72 overflow-y-auto text-xs text-slate-700 list-decimal pl-4">
                    {references.slice(0, 30).map((ref, idx) => (
                      <li key={idx} className="leading-relaxed break-words">
                        {ref}
                      </li>
                    ))}
                  </ol>
                ) : (
                  <p className="text-xs text-slate-400 italic">未识别到参考文献</p>
                )}
                {references.length > 30 && (
                  <p className="text-xs text-slate-400 mt-2">仅展示前 30 条</p>
                )}
              </section>
            </aside>
          </div>
        )}

        {/* ==================== 空闲提示 ==================== */}
        {stage === 'idle' && (
          <section className="bg-white rounded-xl p-8 shadow-sm border border-slate-200 text-center">
            <FileText className="mx-auto text-slate-300 mb-3" size={40} />
            <p className="text-slate-500 text-sm">请先在上方选择文件</p>
            <p className="text-xs text-slate-400 mt-1">支持选择 1-{MAX_BATCH_FILES} 个文件</p>
          </section>
        )}
        {stage === 'preview' && (
          <section className="bg-white rounded-xl p-8 shadow-sm border border-slate-200 text-center">
            <ListChecks className="mx-auto text-blue-300 mb-3" size={40} />
            <p className="text-slate-700 font-medium text-sm">已选择 {pendingFiles.length} 个文件</p>
            <p className="text-xs text-slate-400 mt-1">点击上方「确认解析」按钮开始</p>
          </section>
        )}

        <footer className="mt-6 text-center text-xs text-slate-400">
          MinerU DocParser v1.0.0 · 仅用于本地测试 · Token 由后端安全托管
        </footer>
      </div>
    </div>
  );
}

// 紧凑子阶段进度条
function CompactSubStage({ state, extracted, total }: { state: SubStage | string; extracted: number; total: number }) {
  let displayStage: SubStage | string = state;
  if (state === 'running') {
    displayStage = total > 0 && extracted >= total * 0.8 ? 'inferencing' : 'parsing_page';
  }
  const currentIdx = SUB_STAGE_LIST.findIndex((s) => s.id === displayStage);
  const isTerminal = ['done', 'failed', 'cancelled', 'unknown'].includes(displayStage as string);

  return (
    <div className="flex-1 p-2 bg-slate-50 rounded">
      <div className="text-[10px] text-slate-500 mb-1.5 font-medium">MinerU 处理流程</div>
      <div className="flex items-center">
        {SUB_STAGE_LIST.map((ss, idx) => {
          const Icon = ss.icon;
          const isCurrent = ss.id === displayStage;
          const isDone = currentIdx > idx;
          return (
            <div key={ss.id} className="flex items-center flex-1">
              <div className="flex flex-col items-center">
                <div
                  className={`w-6 h-6 rounded-full flex items-center justify-center border-2 transition-all ${
                    isCurrent ? 'bg-blue-100 border-blue-500 text-blue-600 ring-2 ring-blue-100'
                    : isDone ? 'bg-emerald-100 border-emerald-500 text-emerald-600'
                    : 'bg-white border-slate-200 text-slate-300'
                  }`}
                >
                  {isDone ? <CheckCircle2 size={12} /> : <Icon size={10} />}
                </div>
                <span className={`text-[9px] mt-0.5 font-medium whitespace-nowrap ${isCurrent ? 'text-blue-600' : isDone ? 'text-emerald-600' : 'text-slate-400'}`}>
                  {ss.label}
                </span>
              </div>
              {idx < SUB_STAGE_LIST.length - 1 && (
                <div className={`h-0.5 flex-1 mx-0.5 transition-all ${currentIdx > idx ? 'bg-emerald-400' : 'bg-slate-200'}`} />
              )}
            </div>
          );
        })}
      </div>
      {isTerminal && (
        <div className="text-[10px] text-slate-500 mt-1 text-center">
          {displayStage === 'done' && '✓ 已完成'}
          {displayStage === 'failed' && '✗ 失败'}
          {displayStage === 'cancelled' && '⛔ 已取消'}
          {displayStage === 'unknown' && '等待响应'}
        </div>
      )}
    </div>
  );
}