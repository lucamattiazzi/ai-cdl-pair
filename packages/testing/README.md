# @ai-cdl/testing

Synthetic virtual workbook, scripted agents, trace collector, assertions, and reusable adapter
contract tests for AI-CDL integrations.

The endpoint CLI accepts a configurable per-case timeout:

```bash
ai-cdl-contract --endpoint https://agent.example.com/turns --timeout-ms 60000
```

Continuation cases first read the endpoint's emitted `tool_call.id` and echo that exact ID in the
following `tool_result`, so the suite works with stateful model-backed bridges as well as scripted
adapters. Each case uses a distinct turn ID.
