export {};

declare global {
  interface Window {
    refreshPCFAttachments?: () => Promise<void>;
  }
}