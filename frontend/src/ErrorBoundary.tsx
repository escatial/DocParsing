import { Component, type ReactNode } from 'react';
import { AlertTriangle, RotateCw } from 'lucide-react';

interface Props {
  children: ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

export class ErrorBoundary extends Component<Props, State> {
  state: State = { hasError: false, error: null };

  static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, info: React.ErrorInfo) {
    console.error('[ErrorBoundary] 捕获到 React 错误:', error, info);
  }

  handleReload = () => {
    window.location.reload();
  };

  handleReset = () => {
    this.setState({ hasError: false, error: null });
  };

  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-slate-50 to-blue-50 p-6">
          <div className="bg-white rounded-xl shadow-lg border border-slate-200 max-w-lg w-full p-6">
            <div className="flex items-center gap-2 text-red-600 mb-3">
              <AlertTriangle size={24} />
              <h1 className="text-xl font-bold">页面出现错误</h1>
            </div>
            <p className="text-sm text-slate-600 mb-3">
              应用在渲染过程中抛出了异常。已阻止崩溃扩散，但请刷新页面或重试。
            </p>
            {this.state.error && (
              <pre className="bg-slate-100 text-xs text-slate-800 p-3 rounded overflow-auto max-h-40 mb-4">
                {this.state.error.message}
                {'\n'}
                {this.state.error.stack?.split('\n').slice(0, 5).join('\n')}
              </pre>
            )}
            <div className="flex gap-2">
              <button
                onClick={this.handleReload}
                className="flex-1 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium py-2 px-4 rounded-lg flex items-center justify-center gap-1.5"
              >
                <RotateCw size={14} /> 刷新页面
              </button>
              <button
                onClick={this.handleReset}
                className="text-sm text-slate-600 hover:text-slate-800 border border-slate-300 rounded-lg px-3"
              >
                重试
              </button>
            </div>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}