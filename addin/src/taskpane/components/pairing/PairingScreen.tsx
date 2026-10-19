import { useEffect, useState } from 'react';
import { pair, SIDECAR_BASE } from '../../../connection.js';

export interface PairingScreenProps {
  onPaired: () => void;
}

export function PairingScreen({ onPaired }: PairingScreenProps): JSX.Element {
  const [code, setCode] = useState('');
  const [error, setError] = useState('');
  const [sidecarUp, setSidecarUp] = useState<boolean | null>(null);

  // Poll health so the user can see the sidecar come alive while pairing.
  useEffect(() => {
    let stop = false;
    const check = async (): Promise<void> => {
      try {
        const res = await fetch(`${SIDECAR_BASE}/api/health`);
        if (!stop) setSidecarUp(res.ok);
      } catch {
        if (!stop) setSidecarUp(false);
      }
    };
    void check();
    const timer = setInterval(() => void check(), 3000);
    return () => {
      stop = true;
      clearInterval(timer);
    };
  }, []);

  return (
    <div className="app center">
      <h2>Pair with the companion app</h2>
      <p className="hint">
        {sidecarUp === true
          ? '✓ Sidecar detected. Enter the pairing code from its console.'
          : sidecarUp === false
            ? 'Sidecar not detected — start it with: pnpm dev:sidecar'
            : 'Looking for the sidecar…'}
      </p>
      <input
        value={code}
        maxLength={6}
        placeholder="ABC123"
        aria-label="Pairing code"
        onChange={(e) => setCode(e.target.value.toUpperCase())}
      />
      <button
        onClick={() => {
          void pair(code).then(onPaired, (e: unknown) =>
            setError(e instanceof Error ? e.message : String(e)),
          );
        }}
        disabled={code.length !== 6}
      >
        Pair
      </button>
      {error && <p className="bubble error">{error}</p>}
    </div>
  );
}
