import { z } from 'zod';
import type { ChangeOp, ChangePreviewItem, ChangeSetStatus, UndoFidelity } from './changeset.js';

/**
 * WebSocket protocol between the add-in (browser) and the sidecar.
 *
 * The add-in is the workbook executor ("hands"): the sidecar sends it tool
 * execution requests and it answers. The add-in is also the UI: the sidecar
 * streams agent/chat events to it.
 *
 * Every message carries a protocol version; both sides refuse mismatched
 * majors so silent version skew is impossible.
 */

export const PROTOCOL_VERSION = 1;

export interface Envelope<T extends string, P> {
  v: number;
  id: string;
  type: T;
  payload: P;
}

/* ---------- add-in -> sidecar ---------- */

/**
 * Client frames are untrusted input: the sidecar parses every one against
 * these schemas before dispatch, and the static types are derived from them.
 */
const id = z.string().min(1).max(128);

function envelope<T extends string, P extends z.ZodTypeAny>(type: T, payload: P) {
  return z.object({ v: z.number(), id, type: z.literal(type), payload });
}

const clientHelloSchema = envelope(
  'hello',
  z.object({
    token: z.string().max(512),
    workbookId: z.string().min(1).max(512),
    workbookName: z.string().max(1024),
    client: z.enum(['addin', 'devtool']),
  }),
);
const chatRequestSchema = envelope('chat', z.object({ message: z.string() }));
const toolResultSchema = envelope(
  'tool_result',
  z.object({ requestId: id, ok: z.boolean(), result: z.unknown().optional(), error: z.string().optional() }),
);
const approveChangeSetSchema = envelope('approve_changeset', z.object({ changeSetId: id }));
const rejectChangeSetSchema = envelope('reject_changeset', z.object({ changeSetId: id }));
const undoChangeSetSchema = envelope('undo_changeset', z.object({ changeSetId: id }));
const abortRequestSchema = envelope('abort', z.object({ reason: z.string().max(1024).optional() }));

export const clientMessageSchema = z.discriminatedUnion('type', [
  clientHelloSchema,
  chatRequestSchema,
  toolResultSchema,
  approveChangeSetSchema,
  rejectChangeSetSchema,
  undoChangeSetSchema,
  abortRequestSchema,
]);

export type ClientHello = z.infer<typeof clientHelloSchema>;
export type ChatRequest = z.infer<typeof chatRequestSchema>;
export type ToolResult = z.infer<typeof toolResultSchema>;
export type ApproveChangeSet = z.infer<typeof approveChangeSetSchema>;
export type RejectChangeSet = z.infer<typeof rejectChangeSetSchema>;
export type UndoChangeSet = z.infer<typeof undoChangeSetSchema>;
export type AbortRequest = z.infer<typeof abortRequestSchema>;
export type ClientMessage = z.infer<typeof clientMessageSchema>;

/* ---------- sidecar -> add-in ---------- */

export type ServerHello = Envelope<
  'hello_ack',
  { ok: boolean; error?: string; sidecarVersion: string }
>;

/** Sidecar asks the add-in to execute a (read) tool or apply ops against Excel. */
export type ToolExecRequest = Envelope<
  'tool_exec',
  { requestId: string; tool: string; args: unknown }
>;

export type ApplyOpsRequest = Envelope<
  'apply_ops',
  { requestId: string; ops: ChangeOp[] }
>;

export type AgentEvent = Envelope<
  'agent_event',
  | { kind: 'thinking_delta'; text: string }
  | { kind: 'text_delta'; text: string }
  | { kind: 'thinking'; text: string; durationMs?: number }
  | { kind: 'text'; text: string }
  | { kind: 'tool_use'; tool: string; summary: string }
  | {
      kind: 'changeset_staged';
      changeSetId: string;
      preview: ChangePreviewItem[];
      undoFidelity: UndoFidelity;
    }
  /** `unknown`: the sidecar no longer has this set (e.g. it restarted). */
  | { kind: 'changeset_status'; changeSetId: string; status: ChangeSetStatus | 'unknown' }
  | { kind: 'error'; message: string }
  | { kind: 'stopped'; reason?: string }
  | { kind: 'done' }
>;

/**
 * The sidecar rejected a client frame. `refId` echoes the frame's id when it
 * had one. Repeated violations close the socket (1008).
 */
export type ProtocolError = Envelope<
  'protocol_error',
  { message: string; issues: string[]; refId?: string }
>;

export type ServerMessage = ServerHello | ToolExecRequest | ApplyOpsRequest | AgentEvent | ProtocolError;

export function makeEnvelope<T extends string, P>(type: T, payload: P, id?: string): Envelope<T, P> {
  return {
    v: PROTOCOL_VERSION,
    id: id ?? globalThis.crypto.randomUUID(),
    type,
    payload,
  };
}
