import { useState } from 'react';
import type { LicenseView } from '../../electron/api.ts';

interface Props {
  license: LicenseView;
  onChange: (l: LicenseView) => void;
  onClose: () => void;
}

export function LicensePanel({ license, onChange, onClose }: Props): React.JSX.Element {
  const [key, setKey] = useState('');
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState<string | null>(license.problem ?? null);

  const activate = async (): Promise<void> => {
    setBusy(true);
    try {
      const next = await window.deckforge.license.activate(key.trim());
      onChange(next);
      setMessage(next.problem ?? (next.mode === 'licensed' ? 'Thank you! DeckForge is activated.' : null));
      if (next.mode === 'licensed') setKey('');
    } catch (e) {
      setMessage((e as Error).message);
    } finally {
      setBusy(false);
    }
  };
  const deactivate = async (): Promise<void> => {
    onChange(await window.deckforge.license.deactivate());
    setMessage('The license was removed from this computer.');
  };

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" role="dialog" aria-label="License" onClick={(e) => e.stopPropagation()}>
        <h2>License</h2>
        <p>{license.summary}</p>
        {license.mode === 'licensed' && (
          <dl className="facts">
            {license.licensee && (<><dt>Licensed to</dt><dd>{license.licensee}</dd></>)}
            {license.seats && (<><dt>Seats</dt><dd>{license.seats}</dd></>)}
            {license.updatesUntil && (<><dt>Free updates until</dt><dd>{license.updatesUntil}</dd></>)}
          </dl>
        )}
        {license.mode !== 'licensed' && (
          <>
            <label className="block">License key
              <textarea rows={4} value={key} onChange={(e) => setKey(e.target.value)} placeholder="Paste the key from your purchase email, or the contents of your license file" />
            </label>
            <div className="row">
              <button className="primary" disabled={!key.trim() || busy} onClick={() => void activate()}>{busy ? 'Checking…' : 'Activate'}</button>
            </div>
          </>
        )}
        {message && <p className="muted">{message}</p>}
        <div className="row end">
          {license.mode === 'licensed' && <button onClick={() => void deactivate()}>Remove from this computer</button>}
          <button onClick={onClose}>Close</button>
        </div>
      </div>
    </div>
  );
}
