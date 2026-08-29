import type {
  CollaborationAdapter,
  CollaborationConnection,
  CollaborationEnvelope,
  CollaborationPresence,
  CollaborationPresenceMessage
} from "../domain/types";

export type CollaborationClientProtocolMessage =
  | { type: "join"; protocolVersion: 1; workbookId: string; clientId: string }
  | { type: "leave"; protocolVersion: 1; workbookId: string; clientId: string }
  | { type: "operations"; protocolVersion: 1; envelope: CollaborationEnvelope }
  | CollaborationPresenceMessage;

export type CollaborationServerProtocolMessage =
  | {
      type: "ready";
      protocolVersion: 1;
      workbookId: string;
      envelopes?: CollaborationEnvelope[];
      presences?: CollaborationPresence[];
    }
  | { type: "operations"; protocolVersion: 1; envelope: CollaborationEnvelope }
  | CollaborationPresenceMessage
  | { type: "error"; protocolVersion: 1; workbookId?: string; code: string; message: string };

export interface CollaborationTransportCallbacks {
  receive(message: CollaborationServerProtocolMessage): void;
  error(error: unknown): void;
}

export interface CollaborationTransport {
  connect(callbacks: CollaborationTransportCallbacks): void | Promise<void>;
  send(message: CollaborationClientProtocolMessage): void | Promise<void>;
  disconnect?(): void | Promise<void>;
}

export const createCollaborationTransportAdapter = (transport: CollaborationTransport): CollaborationAdapter => {
  let connection: CollaborationConnection | undefined;
  let connected = false;
  const pendingMessages: CollaborationClientProtocolMessage[] = [];
  const presences = new Map<string, CollaborationPresence>();

  const reportError = (error: unknown): void => {
    connection?.receiveError?.(error);
  };

  const applyPresence = (message: CollaborationPresenceMessage): void => {
    if (message.type === "presence:update") {
      presences.set(message.clientId, message.presence);
    } else {
      presences.delete(message.clientId);
    }
    connection?.receivePresence?.(message);
  };

  const receive = (message: CollaborationServerProtocolMessage): void => {
    if (!connection) {
      return;
    }
    if (message.type === "error") {
      reportError(new Error(`${message.code}: ${message.message}`));
      return;
    }
    if (message.type === "ready") {
      if (message.workbookId !== connection.workbookId) {
        reportError(new Error("Collaboration ready message does not match this workbook."));
        return;
      }
      for (const presence of message.presences ?? []) {
        presences.set(presence.clientId, presence);
      }
      for (const envelope of message.envelopes ?? []) {
        connection.receive(envelope);
      }
      for (const presence of message.presences ?? []) {
        connection.receivePresence?.({
          type: "presence:update",
          workbookId: connection.workbookId,
          clientId: presence.clientId,
          sequence: presence.sequence,
          timestamp: presence.updatedAt,
          presence
        });
      }
      return;
    }
    if (message.type === "operations") {
      connection.receive(message.envelope);
      return;
    }
    applyPresence(message);
  };

  const sendOrQueue = async (message: CollaborationClientProtocolMessage): Promise<void> => {
    if (!connected) {
      pendingMessages.push(message);
      return;
    }
    await transport.send(message);
  };

  return {
    async connect(nextConnection) {
      connection = nextConnection;
      await transport.connect({ receive, error: reportError });
      connected = true;
      await transport.send({
        type: "join",
        protocolVersion: 1,
        workbookId: nextConnection.workbookId,
        clientId: nextConnection.clientId
      });
      for (const message of pendingMessages.splice(0)) {
        await transport.send(message);
      }
    },
    send(envelope) {
      return sendOrQueue({ type: "operations", protocolVersion: 1, envelope });
    },
    getPresence(workbookId) {
      return [...presences.values()].filter((presence) => connection?.workbookId === workbookId && presence.expiresAt > Date.now());
    },
    updatePresence(message) {
      presences.set(message.clientId, message.presence);
      return sendOrQueue(message);
    },
    removePresence(message) {
      presences.delete(message.clientId);
      return sendOrQueue(message);
    },
    async disconnect() {
      const currentConnection = connection;
      if (connected && currentConnection) {
        await transport.send({
          type: "leave",
          protocolVersion: 1,
          workbookId: currentConnection.workbookId,
          clientId: currentConnection.clientId
        });
      }
      connected = false;
      connection = undefined;
      pendingMessages.length = 0;
      presences.clear();
      await transport.disconnect?.();
    }
  };
};