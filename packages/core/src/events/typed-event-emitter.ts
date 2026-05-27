type Listener<T> = (payload: T) => void;

export class TypedEventEmitter<TEvents extends object> {
  private readonly listeners = new Map<keyof TEvents, Set<Listener<unknown>>>();

  on<TKey extends keyof TEvents>(
    eventName: TKey,
    listener: Listener<TEvents[TKey]>
  ): () => void {
    const listeners = this.listeners.get(eventName) ?? new Set<Listener<unknown>>();
    listeners.add(listener as Listener<unknown>);
    this.listeners.set(eventName, listeners);

    return () => {
      listeners.delete(listener as Listener<unknown>);
      if (listeners.size === 0) {
        this.listeners.delete(eventName);
      }
    };
  }

  emit<TKey extends keyof TEvents>(eventName: TKey, payload: TEvents[TKey]): void {
    const listeners = this.listeners.get(eventName);
    if (!listeners) {
      return;
    }

    for (const listener of listeners) {
      try {
        (listener as Listener<TEvents[TKey]>)(payload);
      } catch {
        continue;
      }
    }
  }

  clear(): void {
    this.listeners.clear();
  }
}