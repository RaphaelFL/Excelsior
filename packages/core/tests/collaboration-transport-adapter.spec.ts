import { describe, expect, it, vi } from "vitest";
import {
  createCollaborationTransportAdapter,
  type CollaborationClientProtocolMessage,
  type CollaborationConnection,
  type CollaborationServerProtocolMessage,
  type CollaborationTransportCallbacks
} from "../src";

describe("collaboration transport adapter", () => {
  it("joins, queues operations, replays state and leaves through a generic transport", async () => {
    let callbacks: CollaborationTransportCallbacks | undefined;
    let releaseConnection: (() => void) | undefined;
    const sent: CollaborationClientProtocolMessage[] = [];
    const transport = {
      connect: vi.fn((nextCallbacks: CollaborationTransportCallbacks) => {
        callbacks = nextCallbacks;
        return new Promise<void>((resolve) => {
          releaseConnection = resolve;
        });
      }),
      send: vi.fn((message: CollaborationClientProtocolMessage) => {
        sent.push(message);
      }),
      disconnect: vi.fn()
    };
    const adapter = createCollaborationTransportAdapter(transport);
    const received: string[] = [];
    const connection: CollaborationConnection = {
      workbookId: "workbook-1",
      clientId: "client-local",
      receive: (envelope) => received.push(envelope.id),
      receivePresence: (message) => received.push(message.clientId)
    };
    const connecting = Promise.resolve(adapter.connect(connection));
    await adapter.send({
      id: "local:1",
      workbookId: "workbook-1",
      clientId: "client-local",
      sequence: 1,
      timestamp: 1,
      operations: []
    });
    releaseConnection?.();
    await connecting;

    expect(sent.map((message) => message.type)).toEqual(["join", "operations"]);
    callbacks?.receive({
      type: "ready",
      protocolVersion: 1,
      workbookId: "workbook-1",
      envelopes: [{ id: "remote:1", workbookId: "workbook-1", clientId: "remote", sequence: 1, timestamp: 2, operations: [] }],
      presences: [{ clientId: "remote", sequence: 1, updatedAt: 2, expiresAt: Date.now() + 10_000 }]
    } satisfies CollaborationServerProtocolMessage);
    expect(received).toEqual(["remote:1", "remote"]);
    expect(await adapter.getPresence?.("workbook-1")).toHaveLength(1);

    await adapter.disconnect?.();
    expect(sent.at(-1)?.type).toBe("leave");
    expect(transport.disconnect).toHaveBeenCalledOnce();
  });
});