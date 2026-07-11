import { QueryClient } from '@tanstack/react-query';

// 单独导出便于在 main.tsx 包裹根组件
export const queryClient = new QueryClient({
  defaultOptions: {
    queries: {
      retry: 1,
      staleTime: 1000,
    },
  },
});