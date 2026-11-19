/*
 * User data (file names, sheet names, cell texts, slide texts, mail subjects, outline titles …) must never go through
 * Hostwerk's automatic JSX translation – a sheet called "Summary" would otherwise show as "Zusammenfassung".
 * Hostwerk's runtime leaves the props of components marked with Symbol.for('hostwerk.keep') alone:
 *   <Keep>{name}</Keep>                       – text children
 *   <Raw as="span" title={name}>{name}</Raw>  – an element whose own props (title, aria-label, placeholder …) stay as they are
 * Same mechanism as ui/src/shell/common.jsx (no dependency on @hostwerk/kit, so it also works before the kit is loaded).
 */
import { createElement } from 'react';

const KEEP = Symbol.for('hostwerk.keep');

export function Keep({ children }) { return children === undefined ? null : children; }
Keep[KEEP] = true;

export function Raw({ as = 'span', ...props }) { return createElement(as, props); }
Raw[KEEP] = true;
