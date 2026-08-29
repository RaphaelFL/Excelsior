# Base de Colaboração

O Excelsior não inclui servidor, banco, autenticação ou implementação de WebSocket. O pacote `@excelsior/core` fornece somente o contrato necessário para conectar esses componentes no futuro.

## Adapter de transporte

`createCollaborationTransportAdapter()` transforma um transporte genérico em `CollaborationAdapter`, que já é consumido pelo `WorkbookEngine`.

```ts
import {
  WorkbookEngine,
  createCollaborationTransportAdapter,
  type CollaborationTransport
} from "@excelsior/core";

declare const transport: CollaborationTransport;

const adapter = createCollaborationTransportAdapter(transport);
const engine = new WorkbookEngine({
  collaboration: {
    adapter,
    clientId: crypto.randomUUID(),
    presenceTtlMs: 30_000
  }
});
```

O transporte futuro implementa apenas três operações:

```ts
interface CollaborationTransport {
  connect(callbacks: CollaborationTransportCallbacks): void | Promise<void>;
  send(message: CollaborationClientProtocolMessage): void | Promise<void>;
  disconnect?(): void | Promise<void>;
}
```

Pode ser implementado com WebSocket, SignalR, WebTransport, BroadcastChannel ou um transporte de testes em memória.

## Protocolo

Mensagens do cliente:

- `join`: entra na sala do workbook;
- `operations`: envia operações serializáveis;
- `presence:update` e `presence:remove`: atualizam cursor e seleção;
- `leave`: encerra a participação.

Mensagens do servidor:

- `ready`: entrega operações pendentes e presenças atuais;
- `operations`: entrega alterações remotas;
- `presence:update` e `presence:remove`: atualizam participantes;
- `error`: informa falha controlada.

Todas as mensagens de controle usam `protocolVersion: 1`. O adapter enfileira alterações produzidas antes da conexão, envia `join` antes dessa fila, aplica replay inicial sem eco e limpa seu estado ao desconectar.

## Responsabilidades do backend futuro

- autenticar o usuário;
- autorizar acesso ao `workbookId`;
- persistir snapshots e operações;
- ordenar ou rejeitar mensagens inválidas;
- distribuir mensagens entre clientes da mesma sala;
- controlar reconexão, retenção e compactação de histórico;
- aplicar limites de tamanho e frequência.

Essas responsabilidades não fazem parte da biblioteca e podem ser implementadas sem alterar o `WorkbookEngine` ou o renderer.