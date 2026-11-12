/**
 * 对话模拟器的进程内 transport：AgentTransport → streamForProvider('custom', …)，
 * 与 sheets-main.ts 的 aiStream 走完全相同的协议层（openai-compatible SSE /
 * usage chunk / 空流重试 / 工具参数分片），只把 Electron IPC 边界换成 Node fetch。
 */
import type {
  AgentStreamCallbacks,
  AgentStreamHandle,
  AgentStreamRequest,
  AgentTransport,
} from '@genoffice/agent-core'
import { streamForProvider, type AiProviderConfig } from '@genoffice/ai-provider'

export interface SimTransportOptions {
  /** e.g. http://127.0.0.1:8080/v1 */
  baseUrl: string
  model: string
  /** 应用默认 DEFAULT_MAX_OUTPUT_TOKENS=32768（reasoning 与正文共享该预算） */
  maxTokens: number
  /** 端点 max_model_len（210 vLLM: 262144）→ loop 压缩预算校准 */
  contextWindow: number
}

export function createSimTransport(options: SimTransportOptions): AgentTransport {
  const config: AiProviderConfig = { apiKey: 'local', model: options.model, baseUrl: options.baseUrl }
  return {
    contextWindow: () => options.contextWindow,
    stream(request: AgentStreamRequest, callbacks: AgentStreamCallbacks): AgentStreamHandle {
      const controller = new AbortController()
      let settled = false
      const finish = () => {
        if (!settled) {
          settled = true
          callbacks.onDone()
        }
      }
      const fail = (message: string) => {
        if (!settled) {
          settled = true
          callbacks.onError(message)
        }
      }
      void streamForProvider(
        'custom',
        config,
        request.system,
        request.messages,
        request.tools,
        options.maxTokens,
        {
          onDelta: (text) => callbacks.onDelta(text),
          onReasoningDelta: (text) => callbacks.onReasoning?.(text),
          onToolCall: (call) => callbacks.onToolCall(call),
          onStopReason: (reason) => callbacks.onStopReason?.(reason),
          onUsage: (usage) => callbacks.onUsage?.(usage),
          signal: controller.signal,
        },
      )
        .then(finish)
        .catch((e) => {
          // 契约：cancel() 之后 transport 仍要 emit onDone；abort 拒绝 = 取消而非错误
          if (controller.signal.aborted) finish()
          else fail(e instanceof Error ? e.message : String(e))
        })
      return { cancel: () => controller.abort() }
    },
  }
}
