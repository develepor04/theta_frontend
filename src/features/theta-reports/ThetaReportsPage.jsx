import React, { useMemo, useState } from 'react';
import {
  AlertTriangle, CheckCircle2, Clock, Download, Filter, Lightbulb, Loader2, Send,
} from 'lucide-react';
import useStore from '../../store/useStore';
import useSheetData from '../../hooks/useSheetData';
import useIsMobile from '../../hooks/useIsMobile';
import { deriveInsights } from './deriveInsights';
import './thetaReports.css';

const TABS = [
  { id: 'timeline', label: 'Timeline' },
  { id: 'milestones', label: 'Milestones' },
  { id: 'critical', label: 'Critical Path' },
  { id: 'deviations', label: 'Deviations' },
  { id: 'onepager', label: 'One-Page Summary' },
];

/**
 * Theta Sheets workflow "View Reports" surface — matches the Reports & Analysis
 * mockups (Timeline / Milestones / Critical Path / Deviations / One-Pager).
 * Data comes from the saved active Theta Sheet only (not the process engine).
 */
export default function ThetaReportsPage() {
  const { user } = useStore();
  const isMobile = useIsMobile();
  const [tab, setTab] = useState('timeline');
  const { metrics, loading: metricsLoading } = useSheetData({ mode: 'metrics', useActive: true, pollIntervalMs: 4000 });
  const { sheet, loading: sheetLoading } = useSheetData({ mode: 'sheet', useActive: true, pollIntervalMs: 5000 });

  const sheets = sheet?.data?.sheets || [];
  const insights = useMemo(
    () => deriveInsights(sheets, metrics || {}),
    [sheets, metrics],
  );

  const projectName = user?.company_name
    ? `${user.company_name} Programme`
    : 'Theta Sheets Programme';

  if ((metricsLoading || sheetLoading) && !metrics && !sheets.length) {
    return (
      <div className="tr-root">
        <div className="tr-empty" style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8, minHeight: 280 }}>
          <Loader2 size={16} className="spinning" /> Loading report…
        </div>
      </div>
    );
  }

  if (!insights.activities.length) {
    return (
      <div className="tr-root">
        <div className="tr-wrap">
          <Header projectName={projectName} />
          <div className="tr-card tr-empty">
            No schedule activities found in the saved Theta Sheet yet.
            Save a workbook with Activity ID / Activity Name columns, then open View Reports again.
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="tr-root">
      <div className="tr-wrap">
        <Header projectName={projectName} />

        <div className="tr-tabs" role="tablist">
          {TABS.map((t) => (
            <button
              key={t.id}
              type="button"
              role="tab"
              aria-selected={tab === t.id}
              className={`tr-tab${tab === t.id ? ' tr-tab--active' : ''}`}
              onClick={() => setTab(t.id)}
            >
              {t.label}
            </button>
          ))}
        </div>

        {tab === 'timeline' && <TimelineTab phases={insights.phases} />}
        {tab === 'milestones' && <MilestonesTab milestones={insights.milestones} isMobile={isMobile} />}
        {tab === 'critical' && <CriticalPathTab critical={insights.criticalPath} />}
        {tab === 'deviations' && <DeviationsTab deviations={insights.deviations} />}
        {tab === 'onepager' && (
          <OnePagerTab
            onePager={insights.onePager}
            metrics={metrics}
            projectName={projectName}
            company={user?.company_name}
          />
        )}
      </div>
    </div>
  );
}

function Header({ projectName }) {
  return (
    <div className="tr-header">
      <div>
        <h1 className="tr-title">Reports & Analysis</h1>
        <p className="tr-subtitle">
          What changed, what is delayed, and what needs attention across {projectName}
          {' '}
          — timeline, milestones, critical path, and deviations in one place.
        </p>
      </div>
      <div className="tr-actions">
        <button type="button" className="tr-btn">
          <Filter size={15} /> Filters
        </button>
        <button type="button" className="tr-btn tr-btn--primary">
          <Download size={15} /> Export one-pager
        </button>
      </div>
    </div>
  );
}

function TimelineTab({ phases }) {
  const axis = buildMonthAxis(phases.totalMonths);
  return (
    <div className="tr-card">
      <div className="tr-card-head">
        <div>
          <h2 className="tr-card-title">Project phase timeline</h2>
          <div className="tr-card-meta">
            {phases.totalMonths}-month programme · month {phases.monthNow} of {phases.totalMonths}
          </div>
        </div>
        {phases.baselineDeltaDays > 0 && (
          <span className="tr-badge">
            −{phases.baselineDeltaDays} days vs baseline
          </span>
        )}
      </div>

      {phases.rows.map((row) => (
        <div key={row.phase} className="tr-phase-row">
          <div className="tr-phase-name">{row.phase}</div>
          <div className="tr-track">
            <div
              className={`tr-bar tr-bar--${row.tone}`}
              style={{ left: `${row.left}%`, width: `${row.width}%` }}
            />
          </div>
          <div className="tr-phase-pct">{row.pct}%</div>
        </div>
      ))}

      <div className="tr-axis">
        {axis.map((label) => <span key={label}>{label}</span>)}
      </div>

      <div className="tr-legend">
        <span><i className="tr-dot" style={{ background: '#16a34a' }} /> On track</span>
        <span><i className="tr-dot" style={{ background: '#d97706' }} /> At risk</span>
        <span><i className="tr-dot" style={{ background: '#dc2626' }} /> Delayed</span>
        <span><i className="tr-dot" style={{ background: '#cbd5e1' }} /> Not started</span>
      </div>
    </div>
  );
}

function MilestonesTab({ milestones, isMobile }) {
  const cards = milestones.cards || [];
  const nowPct = 42;
  return (
    <div className="tr-card">
      <div className="tr-card-head">
        <div>
          <h2 className="tr-card-title">Milestone schedule</h2>
          <div className="tr-card-meta">target dates across the delivery window · status by colour</div>
        </div>
        <span className="tr-badge" style={{ background: '#fff', border: '1px solid #e5e7eb', color: '#334155' }}>
          <i className="tr-dot" style={{ background: '#eab308' }} /> {milestones.atRisk} at risk
          {' · '}
          <i className="tr-dot" style={{ background: '#dc2626' }} /> {milestones.delayed} delayed
        </span>
      </div>

      <div className="tr-ms-track-wrap" style={{ '--tr-now': `${nowPct}%` }}>
        <div className="tr-ms-line" />
        <div className="tr-ms-now">NOW</div>
        <div className="tr-ms-nodes">
          {cards.map((m, i) => {
            const left = 8 + (i / Math.max(cards.length - 1, 1)) * 84;
            return (
              <div key={m.id || i} className="tr-ms-node" style={{ left: `${left}%` }}>
                <div className={`tr-diamond tr-diamond--${m.tone}`} />
                {!isMobile && (
                  <>
                    <div className="tr-ms-node-title">{truncate(m.title, 28)}</div>
                    <div className="tr-ms-node-date">{fmtDate(m.date)}</div>
                  </>
                )}
              </div>
            );
          })}
        </div>
      </div>

      <div className="tr-legend" style={{ marginBottom: 8 }}>
        <span><i className="tr-dot" style={{ background: '#16a34a' }} /> Achieved</span>
        <span><i className="tr-dot" style={{ background: '#eab308' }} /> At risk</span>
        <span><i className="tr-dot" style={{ background: '#dc2626' }} /> Delayed</span>
        <span><i className="tr-dot" style={{ background: '#94a3b8' }} /> Scheduled</span>
      </div>

      <div className="tr-ms-grid">
        {cards.map((m) => (
          <div key={`card-${m.id}`} className="tr-ms-card">
            <div className="tr-ms-card-top">
              <span className={`tr-pill tr-pill--${m.tone}`}>
                <i className="tr-dot" style={{ background: 'currentColor' }} /> {m.status}
              </span>
              <span style={{ fontSize: 12, color: '#94a3b8' }}>{fmtDate(m.date)}</span>
            </div>
            <h3 className="tr-ms-card-title">{m.title}</h3>
            <p className="tr-ms-card-detail">{m.detail}</p>
          </div>
        ))}
      </div>
    </div>
  );
}

function CriticalPathTab({ critical }) {
  return (
    <>
      <div className="tr-insight">
        <Lightbulb size={18} color="#d97706" style={{ flexShrink: 0, marginTop: 1 }} />
        <div>{critical.insight}</div>
      </div>

      <div className="tr-card" style={{ marginBottom: 14 }}>
        <div style={{ overflowX: 'auto' }}>
          <table className="tr-table">
            <thead>
              <tr>
                <th>ID</th>
                <th>Critical activity</th>
                <th>Duration</th>
                <th>Float</th>
                <th>Status</th>
                <th>Impact if delayed</th>
              </tr>
            </thead>
            <tbody>
              {critical.rows.map((r) => (
                <tr key={r.id}>
                  <td style={{ color: '#94a3b8' }}>{r.id}</td>
                  <td style={{ fontWeight: 700, color: '#0f172a' }}>{r.name}</td>
                  <td>{r.duration}d</td>
                  <td className={r.floatDays === 0 ? 'tr-float-zero' : ''}>{r.floatDays}d</td>
                  <td>
                    <span className={`tr-pill tr-pill--${r.tone}`}>
                      <i className="tr-dot" style={{ background: 'currentColor' }} /> {r.status}
                    </span>
                  </td>
                  <td>{r.impactIfDelayed}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="tr-card">
        <div className="tr-card-head">
          <div>
            <h2 className="tr-card-title">Critical chain</h2>
            <div className="tr-card-meta">longest path to completion</div>
          </div>
        </div>
        <div className="tr-chain">
          {(critical.chain.length ? critical.chain : ['No chain']).map((node, i, arr) => (
            <React.Fragment key={`${node}-${i}`}>
              <div className={`tr-chain-node${i >= Math.max(arr.length - 3, 1) ? ' tr-chain-node--future' : ''}`}>
                {node}
              </div>
              {i < arr.length - 1 && <span className="tr-chain-arrow">→</span>}
            </React.Fragment>
          ))}
        </div>
      </div>
    </>
  );
}

function DeviationsTab({ deviations }) {
  const { kpis, rows, delayCategories } = deviations;
  return (
    <>
      <div className="tr-kpi-row">
        <div className="tr-kpi">
          <div className="tr-kpi-label">Total deviations</div>
          <div className="tr-kpi-value">{kpis.total}</div>
          <div className="tr-kpi-sub">This reporting period</div>
          <AlertTriangle size={16} color="#d97706" className="tr-kpi-icon" />
        </div>
        <div className="tr-kpi">
          <div className="tr-kpi-label">Schedule-related</div>
          <div className="tr-kpi-value" style={{ color: '#dc2626' }}>{kpis.scheduleRelated}</div>
          <Clock size={16} color="#dc2626" className="tr-kpi-icon" />
        </div>
        <div className="tr-kpi">
          <div className="tr-kpi-label">Cost-related</div>
          <div className="tr-kpi-value">{kpis.costRelated}</div>
          <span className="tr-kpi-icon" style={{ fontWeight: 800, color: '#d97706' }}>$</span>
        </div>
        <div className="tr-kpi">
          <div className="tr-kpi-label">Resolved</div>
          <div className="tr-kpi-value" style={{ color: '#16a34a' }}>{kpis.resolved}</div>
          <div className="tr-kpi-sub" style={{ color: '#16a34a' }}>↑ {kpis.closureRate}% closure rate</div>
          <CheckCircle2 size={16} color="#16a34a" className="tr-kpi-icon" />
        </div>
      </div>

      <div className="tr-card" style={{ marginBottom: 14 }}>
        <div style={{ overflowX: 'auto' }}>
          <table className="tr-table">
            <thead>
              <tr>
                <th>Activity</th>
                <th>Category</th>
                <th>Planned</th>
                <th>Actual</th>
                <th>Variance</th>
                <th>RAG</th>
                <th>Owner</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={`${r.activity}-${i}`} style={i === rows.length - 1 ? { background: '#111827', color: '#fff' } : undefined}>
                  <td style={{ fontWeight: 650, color: i === rows.length - 1 ? '#fff' : '#0f172a' }}>{r.activity}</td>
                  <td><span className="tr-pill tr-pill--scheduled">{r.category}</span></td>
                  <td>{r.planned}</td>
                  <td>{r.actual}</td>
                  <td style={{ color: i === rows.length - 1 ? '#fca5a5' : '#dc2626', fontWeight: 700 }}>{r.variance}</td>
                  <td><span className={`tr-pill tr-pill--${r.rag}`}>{r.rag}</span></td>
                  <td>{r.owner}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="tr-card">
        <div className="tr-card-head">
          <div>
            <h2 className="tr-card-title">Delay categories</h2>
            <div className="tr-card-meta">root cause breakdown</div>
          </div>
        </div>
        {delayCategories.map((c) => (
          <div key={c.key} className="tr-cause-row">
            <div>{c.label}</div>
            <div className="tr-cause-bar">
              <div className="tr-cause-fill" style={{ width: `${c.pct}%`, background: c.color }} />
            </div>
            <div style={{ fontWeight: 700, textAlign: 'right' }}>{c.pct}%</div>
          </div>
        ))}
        <div className="tr-legend">
          {delayCategories.map((c) => (
            <span key={`leg-${c.key}`}>
              <i className="tr-dot" style={{ background: c.color }} /> {c.label.split(' ')[0]}
            </span>
          ))}
        </div>
      </div>
    </>
  );
}

function OnePagerTab({ onePager, metrics, projectName, company }) {
  const costLabel = formatCost(onePager.costExposure, metrics);
  return (
    <div className="tr-card">
      <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'flex-start' }}>
        <div>
          <div className="tr-onepager-label">Executive one-pager</div>
          <h2 className="tr-onepager-title">{projectName}</h2>
          <p className="tr-onepager-meta">
            {company || 'Theta'}
            {' · '}
            Reporting period ending {onePager.reportingDate.toLocaleDateString([], { day: '2-digit', month: 'short', year: 'numeric' })}
          </p>
        </div>
        <span className={`tr-pill tr-pill--${statusTone(onePager.healthStatus)}`}>
          <i className="tr-dot" style={{ background: 'currentColor' }} /> {onePager.healthStatus}
        </span>
      </div>

      <div className="tr-metric-strip">
        <div>
          <div className="tr-kpi-label">Progress</div>
          <div className="tr-kpi-value" style={{ color: '#16a34a' }}>
            {onePager.progress != null ? `${onePager.progress}%` : '—'}
          </div>
        </div>
        <div>
          <div className="tr-kpi-label">Schedule</div>
          <div className="tr-kpi-value" style={{ color: '#dc2626' }}>
            {onePager.scheduleDays}d
          </div>
        </div>
        <div>
          <div className="tr-kpi-label">Cost exposure</div>
          <div className="tr-kpi-value" style={{ color: '#d97706' }}>{costLabel}</div>
        </div>
        <div>
          <div className="tr-kpi-label">Risk score</div>
          <div className="tr-kpi-value" style={{ color: '#dc2626' }}>{onePager.riskScore}/100</div>
        </div>
      </div>

      <h3 className="tr-section-title">Headline</h3>
      <p className="tr-section-body">{onePager.headline}</p>

      <h3 className="tr-section-title">What needs a decision</h3>
      {onePager.decisions.map((d, i) => (
        <div key={i} className="tr-decision">
          <span className="tr-decision-num">{i + 1}</span>
          <span>{d}</span>
        </div>
      ))}

      <div style={{ display: 'flex', gap: 10, marginTop: 22, flexWrap: 'wrap' }}>
        <button type="button" className="tr-btn tr-btn--primary">
          <Download size={15} /> Export as PDF
        </button>
        <button type="button" className="tr-btn">
          <Send size={15} /> Share with board
        </button>
      </div>
    </div>
  );
}

function buildMonthAxis(totalMonths) {
  const marks = [0, 0.25, 0.5, 0.65, 0.85, 1];
  return marks.map((f, i) => {
    const m = Math.round(totalMonths * f);
    if (i === 3) return `M${m} now`;
    return `M${m}`;
  });
}

function fmtDate(d) {
  if (!d) return '—';
  return d.toLocaleDateString([], { day: '2-digit', month: 'short' });
}

function truncate(s, n) {
  const t = String(s || '');
  return t.length > n ? `${t.slice(0, n - 1)}…` : t;
}

function statusTone(status) {
  const s = String(status || '').toLowerCase();
  if (s.includes('track') || s === 'achieved') return 'onTrack';
  if (s.includes('critical') || s.includes('delay')) return 'delayed';
  if (s.includes('risk')) return 'atRisk';
  return 'scheduled';
}

function formatCost(exposure, metrics) {
  if (exposure == null && metrics?.costExposure == null) return '—';
  const v = exposure != null ? exposure : metrics.costExposure;
  // Backend often returns millions already (e.g. 0.04); also handle raw AED.
  if (Math.abs(v) < 50) return `AED ${Number(v).toFixed(1)}M`;
  if (Math.abs(v) >= 1_000_000) return `AED ${(v / 1_000_000).toFixed(1)}M`;
  if (Math.abs(v) >= 1000) return `AED ${Math.round(v / 1000)}K`;
  return `AED ${Math.round(v)}`;
}
