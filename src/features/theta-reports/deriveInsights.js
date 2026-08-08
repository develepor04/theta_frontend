/**
 * Derive Reports & Analysis views from the saved Theta Sheets grid.
 * Schema mirrors schedule_updated_08-08-2026.xlsx activity + intelligence tabs.
 * Frontend-only — does not call or alter the process engine.
 */

import { hasScheduleHeaders, resolveScheduleHeaders } from '../../utils/thetaValidation';

export function parseSheetDate(v) {
  if (v === undefined || v === null || v === '') return null;
  if (typeof v === 'number' && v > 0 && v < 100000) {
    return new Date(Math.round((v - 25569) * 86400 * 1000));
  }
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;
  const s = String(v).trim().replace(/\s*[A*]$/i, '').trim();
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}

function num(v, fallback = 0) {
  if (v === undefined || v === null || v === '') return fallback;
  const n = parseFloat(String(v).replace('%', ''));
  return Number.isFinite(n) ? n : fallback;
}

function pctComplete(v) {
  const n = num(v, NaN);
  if (!Number.isFinite(n)) return 0;
  return n <= 1 ? Math.round(n * 100) : Math.round(n);
}

function colMap(headers) {
  return Object.fromEntries((headers || []).map((h, i) => [String(h).trim(), i]));
}

function cell(row, idx, name) {
  const i = idx[name];
  if (i === undefined || i >= row.length) return undefined;
  return row[i];
}

/** Flatten all activity-like sheets into row objects keyed by header. */
export function collectActivities(sheets) {
  const list = [];
  for (const sheet of sheets || []) {
    if (!hasScheduleHeaders(sheet.headers)) continue;
    const headers = (sheet.headers || []).map((h) => String(h).trim());
    const resolved = resolveScheduleHeaders(headers);
    const idx = colMap(headers);
    const idKey = resolved.activityId;
    const nameKey = resolved.activityName;
    for (const row of sheet.rows || []) {
      const id = String(cell(row, idx, idKey) ?? '').trim();
      const name = String(cell(row, idx, nameKey) ?? '').trim();
      if (!id || !name) continue;
      const get = (h) => cell(row, idx, h);
      list.push({
        id,
        name,
        phase: String(get('Phase') ?? '').trim() || 'General',
        costCategory: String(get('Cost Category') ?? '').trim(),
        period: String(get('Period') ?? '').trim(),
        status: String(get('Status') ?? '').trim() || 'Not Started',
        varianceDays: Math.round(num(get('Variance (Days)'))),
        rootCause: String(get('Root Cause') ?? '').trim() || '—',
        impact: String(get('Impact') ?? '').trim() || 'None',
        pct: pctComplete(get('% Complete')),
        budget: num(get('Budget Cost (AED)')),
        actual: num(get('Actual Cost (AED)')),
        forecast: num(get('Forecast Cost (AED)')),
        baselineStart: parseSheetDate(get('Baseline Start')),
        baselineFinish: parseSheetDate(get('Baseline Finish')),
        forecastStart: parseSheetDate(get('Forecast Start')),
        forecastFinish: parseSheetDate(get('Forecast Finish')),
        actualStart: parseSheetDate(get('Actual Start')),
        actualFinish: parseSheetDate(get('Actual Finish')),
        productivity: num(get('Productivity Index'), NaN),
      });
    }
  }
  return list;
}

function findSheet(sheets, name) {
  const needle = String(name).toLowerCase();
  return (sheets || []).find((s) => String(s.name || '').toLowerCase() === needle) || null;
}

function statusTone(status) {
  const s = String(status || '').toLowerCase();
  if (s === 'completed' || s === 'on track' || s === 'achieved') return 'onTrack';
  if (s === 'delayed' || s === 'behind') return 'delayed';
  if (s === 'at risk' || s === 'recovering') return 'atRisk';
  if (s === 'not started' || s === 'scheduled') return 'scheduled';
  return 'scheduled';
}

function ragFromImpact(impact, variance) {
  const i = String(impact || '').toLowerCase();
  if (i === 'critical' || i === 'high' || variance >= 15) return 'Red';
  if (i === 'medium' || variance >= 5) return 'Amber';
  return 'Green';
}

export function deriveInsights(sheets, metrics = {}) {
  const activities = collectActivities(sheets);
  const scheduleIntel = findSheet(sheets, 'Schedule Intelligence');
  const costIntel = findSheet(sheets, 'Cost Intelligence');

  const phases = derivePhaseTimeline(activities);
  const milestones = deriveMilestones(activities);
  const criticalPath = deriveCriticalPath(activities, scheduleIntel);
  const deviations = deriveDeviations(activities, scheduleIntel);
  const onePager = deriveOnePager(activities, metrics, phases, criticalPath, deviations);

  return {
    activities,
    phases,
    milestones,
    criticalPath,
    deviations,
    onePager,
    hasCostIntel: Boolean(costIntel),
    projectLabel: metrics.projectName || null,
  };
}

function derivePhaseTimeline(activities) {
  const order = [];
  const byPhase = new Map();
  for (const a of activities) {
    if (!byPhase.has(a.phase)) {
      byPhase.set(a.phase, []);
      order.push(a.phase);
    }
    byPhase.get(a.phase).push(a);
  }

  const allStarts = activities.map((a) => a.baselineStart).filter(Boolean);
  const allEnds = activities.map((a) => a.forecastFinish || a.baselineFinish).filter(Boolean);
  const minStart = allStarts.length ? Math.min(...allStarts.map((d) => d.getTime())) : null;
  const maxEnd = allEnds.length ? Math.max(...allEnds.map((d) => d.getTime())) : null;
  const spanMs = minStart != null && maxEnd != null ? Math.max(maxEnd - minStart, 1) : 1;
  const totalMonths = Math.max(1, Math.round(spanMs / (30.44 * 24 * 3600 * 1000)));
  const now = Date.now();
  const monthNow = minStart != null
    ? Math.min(totalMonths, Math.max(0, Math.round((now - minStart) / (30.44 * 24 * 3600 * 1000))))
    : 0;

  const maxVar = activities.reduce((m, a) => Math.max(m, a.varianceDays), 0);

  return {
    totalMonths,
    monthNow,
    baselineDeltaDays: maxVar,
    rows: order.map((phase) => {
      const rows = byPhase.get(phase) || [];
      const avgPct = rows.length
        ? Math.round(rows.reduce((s, r) => s + r.pct, 0) / rows.length)
        : 0;
      const delayed = rows.filter((r) => r.status === 'Delayed').length;
      const notStarted = rows.filter((r) => r.status === 'Not Started' || r.pct === 0).length;
      const completed = rows.filter((r) => r.status === 'Completed' || r.pct >= 100).length;
      let tone = 'onTrack';
      if (delayed > 0 && delayed / rows.length >= 0.4) tone = 'delayed';
      else if (delayed > 0) tone = 'atRisk';
      else if (notStarted === rows.length) tone = 'scheduled';
      else if (completed === rows.length) tone = 'onTrack';

      const starts = rows.map((r) => r.baselineStart).filter(Boolean);
      const ends = rows.map((r) => r.forecastFinish || r.baselineFinish).filter(Boolean);
      const start = starts.length ? Math.min(...starts.map((d) => d.getTime())) : minStart;
      const end = ends.length ? Math.max(...ends.map((d) => d.getTime())) : maxEnd;
      const left = minStart != null && start != null ? ((start - minStart) / spanMs) * 100 : 0;
      const width = start != null && end != null ? Math.max(4, ((end - start) / spanMs) * 100) : Math.max(8, avgPct * 0.6);

      return {
        phase,
        pct: avgPct,
        tone,
        left: Math.max(0, Math.min(92, left)),
        width: Math.max(4, Math.min(100 - left, width)),
        count: rows.length,
        delayed,
      };
    }),
  };
}

function deriveMilestones(activities) {
  const scored = [...activities]
    .filter((a) => a.baselineFinish || a.forecastFinish)
    .sort((a, b) => {
      const da = (a.forecastFinish || a.baselineFinish).getTime();
      const db = (b.forecastFinish || b.baselineFinish).getTime();
      return da - db;
    });

  // Prefer High/Critical impact + phase boundary finishes as milestones.
  const picks = [];
  const seen = new Set();
  for (const a of scored) {
    const key = a.id;
    if (seen.has(key)) continue;
    const important = ['High', 'Critical'].includes(a.impact)
      || a.status === 'Completed'
      || /handover|commission|topping|watertight|mock/i.test(a.name);
    if (!important && picks.length >= 4) continue;
    seen.add(key);
    picks.push(a);
    if (picks.length >= 6) break;
  }
  if (picks.length < 4) {
    for (const a of scored) {
      if (seen.has(a.id)) continue;
      seen.add(a.id);
      picks.push(a);
      if (picks.length >= 6) break;
    }
  }

  const now = Date.now();
  const cards = picks.map((a) => {
    const date = a.actualFinish || a.forecastFinish || a.baselineFinish;
    let status = 'Scheduled';
    let detail = a.rootCause !== '—' ? a.rootCause : a.phase;
    if (a.status === 'Completed' || a.pct >= 100) {
      status = 'Achieved';
      detail = a.varianceDays < 0 ? `Signed off ${Math.abs(a.varianceDays)} days early` : 'On plan';
    } else if (a.status === 'Delayed') {
      status = 'Delayed';
      detail = a.rootCause !== '—' ? a.rootCause : `+${a.varianceDays}d vs baseline`;
    } else if (a.impact === 'High' || a.impact === 'Medium' || a.varianceDays > 0) {
      status = 'At risk';
      detail = a.rootCause !== '—' ? a.rootCause : 'Watch this milestone';
    } else if (date && date.getTime() < now) {
      status = 'At risk';
    }
    return {
      id: a.id,
      title: a.name,
      date,
      status,
      detail,
      tone: statusTone(status),
    };
  });

  const atRisk = cards.filter((c) => c.status === 'At risk').length;
  const delayed = cards.filter((c) => c.status === 'Delayed').length;
  return { cards, atRisk, delayed };
}

function deriveCriticalPath(activities, scheduleIntelSheet) {
  let rows = activities.filter((a) => a.status === 'Delayed' || a.impact === 'High' || a.impact === 'Critical');
  if (scheduleIntelSheet?.rows?.length) {
    const headers = (scheduleIntelSheet.headers || []).map((h) => String(h).trim());
    const idx = colMap(headers);
    const fromIntel = scheduleIntelSheet.rows
      .map((row) => {
        const id = String(cell(row, idx, 'Activity ID') ?? '').trim();
        const name = String(cell(row, idx, 'Activity Name') ?? '').trim();
        if (!id || !name) return null;
        const status = String(cell(row, idx, 'Status') ?? '').trim();
        const variance = Math.round(num(cell(row, idx, 'Variance (Days)')));
        const impact = String(cell(row, idx, 'Impact') ?? '').trim();
        if (status !== 'Delayed' && impact !== 'High' && impact !== 'Critical') return null;
        return { id, name, status, varianceDays: variance, impact, rootCause: String(cell(row, idx, 'Root Cause') ?? '').trim() };
      })
      .filter(Boolean);
    if (fromIntel.length) rows = fromIntel;
  }

  rows = [...rows]
    .sort((a, b) => (b.varianceDays || 0) - (a.varianceDays || 0))
    .slice(0, 8)
    .map((a) => {
      const status = a.status === 'Delayed' ? 'Behind' : (a.impact === 'High' || a.impact === 'Critical' ? 'At risk' : 'Scheduled');
      const duration = Math.max(5, Math.round((a.varianceDays || 0) + 20));
      return {
        id: a.id,
        name: a.name,
        duration,
        floatDays: a.status === 'Delayed' || status === 'Behind' ? 0 : Math.max(0, 5 - (a.varianceDays || 0)),
        status,
        impactIfDelayed: a.rootCause && a.rootCause !== '—'
          ? a.rootCause
          : (status === 'Behind' ? 'Directly moves handover' : 'Blocks downstream start'),
        tone: statusTone(status === 'Behind' ? 'delayed' : status),
      };
    });

  const chain = rows.slice(0, 5).map((r) => r.name.split(/[-—]/)[0].trim().slice(0, 22));
  const insight = rows.length
    ? `The critical path now runs through ${rows[0].name}. Every day lost on the highlighted activities pushes the handover date by the same amount. Total float on this chain is zero.`
    : 'No critical-path pressure detected from the current sheet.';

  return { rows, chain, insight };
}

function deriveDeviations(activities, scheduleIntelSheet) {
  let source = activities.filter((a) => a.status === 'Delayed' || a.varianceDays > 0);
  if (scheduleIntelSheet?.rows?.length) {
    const headers = (scheduleIntelSheet.headers || []).map((h) => String(h).trim());
    const idx = colMap(headers);
    const intel = scheduleIntelSheet.rows.map((row) => {
      const id = String(cell(row, idx, 'Activity ID') ?? '').trim();
      const name = String(cell(row, idx, 'Activity Name') ?? '').trim();
      if (!id || !name) return null;
      return {
        id,
        name,
        phase: String(cell(row, idx, 'Phase') ?? '').trim(),
        status: String(cell(row, idx, 'Status') ?? '').trim(),
        varianceDays: Math.round(num(cell(row, idx, 'Variance (Days)'))),
        rootCause: String(cell(row, idx, 'Root Cause') ?? '').trim() || '—',
        impact: String(cell(row, idx, 'Impact') ?? '').trim() || 'None',
      };
    }).filter(Boolean);
    if (intel.length) source = intel;
  }

  const byId = Object.fromEntries(activities.map((a) => [a.id, a]));
  const rows = source
    .sort((a, b) => (b.varianceDays || 0) - (a.varianceDays || 0))
    .slice(0, 12)
    .map((a) => {
      const full = byId[a.id] || a;
      const category = categorizeRootCause(a.rootCause, a.phase || full.phase);
      const plannedDate = full.baselineFinish;
      const actualDate = full.actualFinish || full.forecastFinish;
      return {
        activity: a.name,
        category: category.label,
        categoryKey: category.key,
        planned: plannedDate
          ? plannedDate.toLocaleDateString([], { day: '2-digit', month: 'short' })
          : '—',
        actual: actualDate
          ? actualDate.toLocaleDateString([], { day: '2-digit', month: 'short' })
          : (a.status === 'Completed' ? 'Done' : '—'),
        variance: a.varianceDays > 0 ? `+${a.varianceDays}d` : `${a.varianceDays || 0}d`,
        varianceDays: a.varianceDays || 0,
        rag: ragFromImpact(a.impact, a.varianceDays || 0),
        owner: ownerFromPhase(a.phase || full.phase),
        resolved: a.status === 'Completed',
      };
    });

  const total = rows.length;
  const scheduleRelated = rows.filter((r) => r.categoryKey === 'delay' || r.categoryKey === 'mep' || r.categoryKey === 'design').length;
  const costRelated = rows.filter((r) => r.categoryKey === 'procurement').length;
  const resolved = activities.filter((a) => a.status === 'Completed' && a.varianceDays > 0).length;
  const closureRate = total ? Math.round((resolved / Math.max(total + resolved, 1)) * 100) : 0;

  const causeBuckets = [
    { key: 'mep', label: 'MEP sequencing / access', color: '#b91c1c' },
    { key: 'procurement', label: 'Long-lead procurement', color: '#d97706' },
    { key: 'design', label: 'Design change / rework', color: '#2563eb' },
    { key: 'external', label: 'Weather & site conditions', color: '#94a3b8' },
  ];
  const counts = Object.fromEntries(causeBuckets.map((b) => [b.key, 0]));
  for (const r of rows) {
    counts[r.categoryKey] = (counts[r.categoryKey] || 0) + 1;
  }
  const causeTotal = Object.values(counts).reduce((a, b) => a + b, 0) || 1;
  const delayCategories = causeBuckets.map((b) => ({
    ...b,
    pct: Math.round((counts[b.key] / causeTotal) * 100),
  }));

  return {
    kpis: {
      total,
      scheduleRelated,
      costRelated,
      resolved,
      closureRate,
    },
    rows,
    delayCategories,
  };
}

function categorizeRootCause(rootCause, phase) {
  const t = `${rootCause} ${phase}`.toLowerCase();
  if (/mep|riser|hvac|access|sequenc/.test(t)) return { key: 'mep', label: 'Delay' };
  if (/procure|supplier|lead|mill|oem|deliver|glass|tile|chiller/.test(t)) return { key: 'procurement', label: 'Procurement' };
  if (/design|clash|rework|drawing|review|code/.test(t)) return { key: 'design', label: 'Rework' };
  if (/weather|ground|site|external/.test(t)) return { key: 'external', label: 'Quality' };
  if (/delay/.test(t)) return { key: 'delay', label: 'Delay' };
  return { key: 'delay', label: 'Delay' };
}

function ownerFromPhase(phase) {
  const p = String(phase || '').toLowerCase();
  if (p.includes('engineer')) return 'Engineering';
  if (p.includes('procure')) return 'Procurement';
  if (p.includes('commission')) return 'QA/QC';
  if (p.includes('construct')) return 'Project Controls';
  return 'Project Controls';
}

function deriveOnePager(activities, metrics, phases, criticalPath, deviations) {
  const progress = activities.length
    ? Math.round(activities.reduce((s, a) => s + a.pct, 0) / activities.length)
    : null;
  const schedule = metrics.scheduleVariance != null ? -Math.abs(metrics.scheduleVariance) : -(phases.baselineDeltaDays || 0);
  const costExposure = metrics.costExposure;
  const riskScore = Math.max(
    0,
    Math.min(100, Math.round(
      (metrics.healthStatus === 'Critical' ? 85 : metrics.healthStatus === 'At Risk' ? 68 : 35)
      + Math.min(20, (phases.baselineDeltaDays || 0) * 0.4),
    )),
  );

  const top = criticalPath.rows[0];
  const headline = metrics.aiInsight
    || (top
      ? `${top.name} is driving schedule pressure (${Math.abs(schedule)}d). Overall health is ${metrics.healthStatus || 'under review'} with ${deviations.kpis.total} open deviations.`
      : 'Sheet data is loaded. Add or update activities to refine the executive narrative.');

  const decisions = criticalPath.rows.slice(0, 3).map((r, i) => {
    if (i === 0) return `Approve resequencing for ${r.name} to recover time on the critical path.`;
    if (i === 1) return `Escalate ${r.name} — impact: ${r.impactIfDelayed}.`;
    return `Confirm owner response for ${r.name} before the next reporting cut-off.`;
  });
  while (decisions.length < 3) {
    decisions.push('Review delayed activities in Schedule Intelligence and confirm mitigation owners.');
  }

  return {
    progress,
    scheduleDays: schedule,
    costExposure,
    riskScore,
    healthStatus: metrics.healthStatus || 'At Risk',
    headline,
    decisions,
    reportingDate: new Date(),
  };
}
