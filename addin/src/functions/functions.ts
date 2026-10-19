/**
 * =AI() custom-function family. Each call forwards to the sidecar's batch
 * engine, which dedupes, caches and rate-limits — a column fill of identical
 * prompts bills once.
 *
 * Failures surface as Excel #VALUE! errors with a message, never as
 * silently wrong values.
 */

const SIDECAR_BASE = 'http://127.0.0.1:8923';

function isInvocation(obj: unknown): obj is CustomFunctions.CancelableInvocation {
  return typeof obj === 'object' && obj !== null && ('onCanceled' in obj || 'setResult' in obj);
}

async function callFunction(
  body: Record<string, unknown>,
  invocation?: CustomFunctions.CancelableInvocation,
): Promise<string> {
  const token = localStorage.getItem('excelai-token');
  if (!token) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      'Not paired with the sidecar. Open the task pane first.',
    );
  }

  const controller = new AbortController();
  if (invocation) {
    invocation.onCanceled = () => {
      controller.abort();
    };
  }

  try {
    const res = await fetch(`${SIDECAR_BASE}/api/functions/run`, {
      method: 'POST',
      headers: { 'content-type': 'application/json', authorization: `Bearer ${token}` },
      body: JSON.stringify(body),
      signal: controller.signal,
    });
    const data = (await res.json().catch(() => ({}))) as { value?: string; error?: string };
    if (!res.ok || data.value === undefined) {
      throw new CustomFunctions.Error(
        CustomFunctions.ErrorCode.notAvailable,
        data.error ?? `Sidecar error (${res.status})`,
      );
    }
    return data.value;
  } catch (e) {
    if (controller.signal.aborted || (e instanceof Error && e.name === 'AbortError')) {
      throw new CustomFunctions.Error(
        CustomFunctions.ErrorCode.notAvailable,
        'Calculation canceled',
      );
    }
    throw e;
  }
}

export async function ai(
  prompt: string,
  inputOrInvocation?: string | CustomFunctions.CancelableInvocation,
  maybeInvocation?: CustomFunctions.CancelableInvocation,
): Promise<string> {
  let input: string | undefined;
  let invocation: CustomFunctions.CancelableInvocation | undefined;

  if (isInvocation(inputOrInvocation)) {
    invocation = inputOrInvocation;
  } else {
    input = inputOrInvocation;
    invocation = maybeInvocation;
  }

  const body: Record<string, unknown> = { kind: 'ai', prompt };
  if (input !== undefined && input !== null) body.input = String(input);
  return callFunction(body, invocation);
}

export async function aiClassify(
  text: string,
  categories: string,
  invocation?: CustomFunctions.CancelableInvocation,
): Promise<string> {
  return callFunction(
    {
      kind: 'classify',
      prompt: '',
      input: String(text),
      categories: String(categories)
        .split(',')
        .map((c) => c.trim())
        .filter(Boolean),
    },
    isInvocation(invocation) ? invocation : undefined,
  );
}

export async function aiExtract(
  text: string,
  field: string,
  invocation?: CustomFunctions.CancelableInvocation,
): Promise<string> {
  return callFunction(
    { kind: 'extract', prompt: String(field), input: String(text) },
    isInvocation(invocation) ? invocation : undefined,
  );
}

export async function aiTranslate(
  text: string,
  targetLang: string,
  invocation?: CustomFunctions.CancelableInvocation,
): Promise<string> {
  return callFunction(
    { kind: 'translate', prompt: '', input: String(text), targetLang },
    isInvocation(invocation) ? invocation : undefined,
  );
}

/* global CustomFunctions */
// With the shared runtime this module is imported by the task pane page;
// associate only where the custom-functions runtime actually exists.
if (typeof CustomFunctions !== 'undefined') {
  CustomFunctions.associate('AI', ai);
  CustomFunctions.associate('AI_CLASSIFY', aiClassify);
  CustomFunctions.associate('AI_EXTRACT', aiExtract);
  CustomFunctions.associate('AI_TRANSLATE', aiTranslate);
}
