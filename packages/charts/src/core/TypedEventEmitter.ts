type EventHandler<TPayload> = (payload: TPayload) => void;

export class TypedEventEmitter<TMap extends object> {
  private readonly listeners = new Map<keyof TMap & string, Set<EventHandler<unknown>>>();

  on<TKey extends keyof TMap & string>(event: TKey, handler: EventHandler<TMap[TKey]>): void {
    const handlers = this.listeners.get(event) ?? new Set<EventHandler<unknown>>();
    handlers.add(handler as EventHandler<unknown>);
    this.listeners.set(event, handlers);
  }

  off<TKey extends keyof TMap & string>(event: TKey, handler: EventHandler<TMap[TKey]>): void {
    const handlers = this.listeners.get(event);
    if (!handlers) {
      return;
    }
    handlers.delete(handler as EventHandler<unknown>);
    if (handlers.size === 0) {
      this.listeners.delete(event);
    }
  }

  emit<TKey extends keyof TMap & string>(event: TKey, payload: TMap[TKey]): void {
    const handlers = this.listeners.get(event);
    if (!handlers || handlers.size === 0) {
      return;
    }

    for (const handler of handlers) {
      try {
        handler(payload);
      } catch {
        // Listeners are isolated to avoid breaking chart rendering flow.
      }
    }
  }

  clear(): void {
    this.listeners.clear();
  }
}
