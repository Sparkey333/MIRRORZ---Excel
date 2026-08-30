/**
 * The formula inspector.
 *
 * The dependency graph already knows which cell originated an error. Excel makes
 * you find that out by clicking Trace Precedents once per hop and following
 * arrows across sheets, which for a model of any size is a job of minutes. So
 * the banner at the top of this panel is the entire point of the feature: "the
 * problem starts at Sheet1!B4", with a click that takes you there.
 *
 * Under it, the precedent tree is expandable rather than expanded, and ranges
 * stay ranges. Exploding a `SUM(A1:A5000)` into five thousand children is what
 * makes tracer arrows unusable, and repeating it here as a tree would be no
 * better.
 */

import { useMemo, useState } from 'react';
import { isError } from '@mirrorz/core';
import { useApp, useController, useDerived } from '../state/context.js';
import {
  buildInspectorModel,
  describeValue,
  formatAddr,
  type PrecedentNode,
} from '../model/inspector-model.js';
import type { CellAddr } from '@mirrorz/formula';

export function FormulaInspector() {
  const controller = useController();
  const snapshot = useApp();
  const addr = controller.activeAddr();

  const model = useDerived(
    () => buildInspectorModel(controller.explain(addr), (a) => controller.explain(a)),
    [addr.sheet, addr.row, addr.col],
  );

  return (
    <aside className="mz-inspector" aria-label="Formula inspector">
      <header className="mz-panel-header">
        <h2>Inspector</h2>
        <button
          type="button"
          aria-label="Close inspector"
          onClick={() => controller.togglePanel('inspector', false)}
        >
          &#215;
        </button>
      </header>

      <div className="mz-inspector-cell">
        <span className="mz-inspector-addr">{model.label}</span>
        <span className="mz-inspector-value" data-error={isError(model.value)}>
          {model.valueText}
        </span>
      </div>

      {model.formula !== undefined ? (
        <pre className="mz-inspector-formula">={model.formula}</pre>
      ) : (
        <p className="mz-inspector-note">This cell holds a literal value.</p>
      )}

      {model.error ? (
        <div className="mz-inspector-error" role="alert">
          <strong className="mz-inspector-headline">{model.error.headline}</strong>
          {model.error.detail ? <p className="mz-inspector-detail">{model.error.detail}</p> : null}
          {model.error.primary ? (
            <button
              type="button"
              className="mz-inspector-jump"
              onClick={() => controller.goTo(model.error!.primary!)}
            >
              Go to {formatAddr(model.error.primary)}
            </button>
          ) : null}
          {model.error.roots.length > 1 ? (
            <ul className="mz-inspector-roots">
              {model.error.roots.map((root) => (
                <li key={formatAddr(root)}>
                  <button type="button" onClick={() => controller.goTo(root)}>
                    {formatAddr(root)}
                  </button>
                </li>
              ))}
            </ul>
          ) : null}
        </div>
      ) : null}

      <section className="mz-inspector-section">
        <h3>Reads from</h3>
        {model.precedents.children.length === 0 ? (
          <p className="mz-inspector-note">Nothing. This cell does not read any other cell.</p>
        ) : (
          <ul className="mz-precedent-tree">
            {model.precedents.children.map((child) => (
              <PrecedentBranch key={child.key} node={child} depth={0} onJump={(a) => controller.goTo(a)} />
            ))}
          </ul>
        )}
      </section>

      <section className="mz-inspector-section">
        <h3>Read by</h3>
        {model.dependents.length === 0 ? (
          <p className="mz-inspector-note">
            Nothing reads this cell{model.isolated ? ', and it reads nothing: it is an island' : ''}.
          </p>
        ) : (
          <ul className="mz-dependent-list">
            {model.dependents.map((dependent) => (
              <li key={formatAddr(dependent)}>
                <button type="button" onClick={() => controller.goTo(dependent)}>
                  {formatAddr(dependent)}
                </button>
              </li>
            ))}
          </ul>
        )}
      </section>

      <footer className="mz-inspector-footer">
        Version {snapshot.version}, {snapshot.calcMode} calculation
      </footer>
    </aside>
  );
}

function PrecedentBranch({
  node,
  depth,
  onJump,
}: {
  node: PrecedentNode;
  depth: number;
  onJump: (addr: CellAddr) => void;
}) {
  // Errors are expanded by default down to the third level, because when a cell
  // is broken the path to the break is the thing being looked for.
  const [open, setOpen] = useState(node.errored && depth < 2);
  const summary = useMemo(() => (node.value === undefined ? '' : describeValue(node.value)), [node.value]);

  return (
    <li className="mz-precedent" data-kind={node.kind} data-error={node.errored}>
      <div className="mz-precedent-row">
        {node.children.length > 0 ? (
          <button
            type="button"
            className="mz-precedent-twisty"
            aria-expanded={open}
            aria-label={open ? `Collapse ${node.label}` : `Expand ${node.label}`}
            onClick={() => setOpen((o) => !o)}
          >
            {open ? '▾' : '▸'}
          </button>
        ) : (
          <span className="mz-precedent-twisty-spacer" />
        )}
        <button
          type="button"
          className="mz-precedent-label"
          disabled={node.kind === 'range'}
          onClick={() => node.addr && onJump(node.addr)}
        >
          {node.label}
        </button>
        {node.formula !== undefined ? <code className="mz-precedent-formula">={node.formula}</code> : null}
        {summary ? <span className="mz-precedent-value">{summary}</span> : null}
        {node.cyclic ? <span className="mz-precedent-flag">circular</span> : null}
        {node.truncated ? <span className="mz-precedent-flag">more not shown</span> : null}
      </div>
      {open && node.children.length > 0 ? (
        <ul className="mz-precedent-children">
          {node.children.map((child) => (
            <PrecedentBranch key={child.key} node={child} depth={depth + 1} onJump={onJump} />
          ))}
        </ul>
      ) : null}
    </li>
  );
}
