/*
 * Hostwerk's <Switch label> uses `label` only as the screen-reader name – nothing is shown next to the toggle. A
 * switch in a dialog or ribbon needs its text visible: <SwitchRow checked onChange label /> = switch + clickable text.
 */
import { Switch } from '@hostwerk/ui';

const ROW = { display: 'inline-flex', alignItems: 'center', gap: 'var(--sp-2)', minWidth: 0 };
export function SwitchRow({ checked, onChange, label, disabled, className = '', small }) {
  return (
    <div className={className} style={ROW}>
      <Switch checked={checked} onChange={onChange} disabled={disabled} label={label} />
      <span className={small ? 'small' : undefined} style={{ cursor: disabled ? 'default' : 'pointer', opacity: disabled ? 0.55 : 1, userSelect: 'none' }} onClick={() => { if (!disabled) onChange(!checked); }} aria-hidden="true">{label}</span>
    </div>
  );
}
