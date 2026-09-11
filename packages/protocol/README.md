# @ai-cdl/protocol

Provider-independent TypeScript types and Zod runtime schemas for AI-CDL protocol `0.2`.
The package describes Excel observation, action, and interaction methods without depending on
Office.js, a model provider, browser APIs, or `@ai-cdl/core`.

```ts
import { parseProtocolMessage, PROTOCOL_VERSION } from "@ai-cdl/protocol";

const message = parseProtocolMessage({
  protocolVersion: PROTOCOL_VERSION,
  type: "request",
  id: "request-1",
  method: "excel.sheet.list",
  params: {},
});
```
