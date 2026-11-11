import { useEffect, useState } from 'react';
import type { LicenseView, UpdateView } from '../../electron/api.ts';

interface Props {
  license: LicenseView;
  onChange: (l: LicenseView) => void;
  onClose: () => void;
}

export function LicensePanel({ license, onChange, onClose }: Props): React.JSX.Element {
  const api = window.deckforge;
  const [text, setText] = useState('');
  const [online, setOnline] = useState(true);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState<string | null>(license.problem ?? null);
  const [requestCode, setRequestCode] = useState<string | null>(null);
  const [updates, setUpdates] = useState<UpdateView | null>(null);

  useEffect(() => {
    void api.updates.status().then(setUpdates);
  }, [api]);

  const run = async (f: () => Promise<LicenseView>, success?: string): Promise<void> => {
    setBusy(true);
    try {
      const next = await f();
      onChange(next);
      setMessage(next.problem ?? success ?? null);
      if (!next.problem) setText('');
    } catch (e) {
      setMessage((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const licensed = license.mode === 'licensed';
  const hasKey = Boolean(license.licensee);

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" role="dialog" aria-label="License" onClick={(e) => e.stopPropagation()}>
        <h2>License</h2>
        <p>{license.summary}</p>

        {hasKey && (
          <dl className="facts">
            <dt>Licensed to</dt><dd>{license.licensee}</dd>
            {license.seats !== undefined && (<><dt>Seats</dt><dd>{license.seats}</dd></>)}
            {license.updatesUntil && (<><dt>Free updates until</dt><dd>{license.updatesUntil}</dd></>)}
            <dt>This computer</dt><dd>{license.activated ? 'Registered' : 'Not registered (DeckForge works the same)'}</dd>
          </dl>
        )}

        {(!licensed || !license.activated) && (
          <>
            <label className="block">{hasKey ? 'Activation file' : 'License key'}
              <textarea
                rows={4}
                value={text}
                onChange={(e) => setText(e.target.value)}
                placeholder={hasKey ? 'Paste the contents of deckforge-activation.txt' : 'Paste the key from your purchase email, or the contents of your license file'}
              />
            </label>
            {!hasKey && license.canActivateOnline && (
              <label className="check">
                <input type="checkbox" checked={online} onChange={(e) => setOnline(e.target.checked)} />
                Register this computer online (one request to the DeckForge license server)
              </label>
            )}
            <div className="row">
              <button className="primary" disabled={!text.trim() || busy} onClick={() => void run(() => api.license.activate(text, { online }), 'Thank you! DeckForge is activated.')}>
                {busy ? 'Checking…' : hasKey ? 'Load activation file' : 'Activate'}
              </button>
              {license.recoveryUrl && !hasKey && (
                <a className="link" href={license.recoveryUrl} target="_blank" rel="noreferrer">Lost your key?</a>
              )}
            </div>
          </>
        )}

        {hasKey && !license.activated && (
          <div className="offline">
            <h3>Computer without internet?</h3>
            <p className="muted small">
              Copy this request code to a computer with internet, open {license.offlineActivationUrl ?? 'the activation page'}, paste the code and your key, and
              load the activation file it gives you here.
            </p>
            {requestCode ? (
              <input className="mono" readOnly value={requestCode} onFocus={(e) => e.currentTarget.select()} />
            ) : (
              <button onClick={() => void api.license.requestCode().then(setRequestCode, (e: Error) => setMessage(e.message))}>Show request code</button>
            )}
          </div>
        )}

        {message && <p className="muted">{message}</p>}

        {updates && (
          <div className="updates">
            <h3>Updates</h3>
            {updates.state === 'disabled' ? (
              <p className="muted small">{updates.message}</p>
            ) : (
              <>
                <label className="check">
                  <input type="checkbox" checked={updates.auto} onChange={(e) => void api.updates.setAuto(e.target.checked).then(setUpdates)} />
                  Check for updates when DeckForge starts
                </label>
                <div className="row">
                  <button disabled={updates.state === 'checking' || updates.state === 'downloading'} onClick={() => void api.updates.check().then(setUpdates)}>
                    {updates.state === 'checking' ? 'Checking…' : 'Check for updates'}
                  </button>
                  {updates.state === 'ready' && <button className="primary" onClick={() => void api.updates.install()}>Restart to update</button>}
                </div>
                {updates.message && <p className="muted small">{updates.message}</p>}
              </>
            )}
          </div>
        )}

        <div className="row end">
          {hasKey && <button onClick={() => void run(() => api.license.deactivate(), 'The license was removed from this computer.')}>Remove from this computer</button>}
          <button onClick={onClose}>Close</button>
        </div>
      </div>
    </div>
  );
}
