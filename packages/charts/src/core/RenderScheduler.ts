type RenderTask = () => void;

export class RenderScheduler {
  private rafId: number | null = null;
  private queuedTask: RenderTask | null = null;

  schedule(task: RenderTask): void {
    this.queuedTask = task;
    if (this.rafId !== null) {
      return;
    }

    if (typeof requestAnimationFrame === "function") {
      this.rafId = requestAnimationFrame(() => this.flush());
      return;
    }

    this.flush();
  }

  flushNow(task: RenderTask): void {
    this.cancel();
    task();
  }

  destroy(): void {
    this.cancel();
    this.queuedTask = null;
  }

  private flush(): void {
    this.rafId = null;
    const task = this.queuedTask;
    this.queuedTask = null;
    task?.();
  }

  private cancel(): void {
    if (this.rafId !== null && typeof cancelAnimationFrame === "function") {
      cancelAnimationFrame(this.rafId);
    }
    this.rafId = null;
  }
}
