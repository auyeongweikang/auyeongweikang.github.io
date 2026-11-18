(() => {
  const D = window.SITE;
  const $ = (s, r = document) => r.querySelector(s);
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const ideaBy = Object.fromEntries(D.ideas.map((i) => [i.ticker, i]));
  const posts = [...D.posts].sort((a, b) => (b.date + b.time).localeCompare(a.date + a.time));
  const postBy = Object.fromEntries(posts.map((p) => [p.id, p]));

  const state = {
    tab: 'Active',
    sort: { k: 'ticker', dir: 1 },
    qmQ: '', // text typed in the quote monitor's filter box
    ctx: null, // ticker loaded in the command bar
    wire: { q: '', type: 'ALL', ticker: null },
    wb: { ticker: null, sheet: 0, row: null }, // Excel model on the MODEL page: sheet shown, row charted
    ws: 'Main',
    layout: null, // per-window position, size and open state
    font: 1,
    range: 'ALL', // how much price history the chart shows
    pal: { items: [], idx: 0 },
  };
  let booted = false;

  // ---------- formatting ----------
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const fmtDate = (iso) => { const [y, m, d] = iso.split('-'); return `${d} ${MONTHS[+m - 1]} ${y.slice(2)}`; };
  const px = (v) => (v == null ? '—' : v.toFixed(2));
  const pct = (v) => (v == null ? '—' : (v > 0 ? '+' : '') + (v * 100).toFixed(1) + '%');
  const cls = (v) => (v == null ? '' : v > 0 ? 'up' : v < 0 ? 'down' : '');

  // Return since the call and distance to target, signed so that positive is good for the call.
  function calc(i) {
    const mark = i.status === 'closed' ? i.exit : i.last;
    const short = i.stance === 'SHORT';
    let ret = null, upside = null;
    if (i.entry != null && mark != null) ret = (short ? i.entry - mark : mark - i.entry) / i.entry;
    if (i.target != null && i.status === 'active') upside = (short ? i.last - i.target : i.target - i.last) / i.last;
    // A position held through options carries its own return, worked out by build_site.py from its legs.
    if (i.position) ret = i.position.ret;
    return { mark, ret, upside };
  }
  // In the monitor an options position shows what the position cost and is worth (per share),
  // not the stock's prices, and no stock target.
  const qmEntry = (i) => (i.position ? i.position.cost : i.entry);
  const qmMark = (i) => (i.position ? i.position.value : calc(i).mark);
  const qmTarget = (i) => (i.position ? null : i.target);
  const qmUpside = (i) => (i.position ? null : calc(i).upside);
  // A leg of an options position that has been closed or has expired gets a row of its own in
  // the monitor's Closed legs tab while the position as a whole is still open. These rows are not
  // ideas: they stay out of D.ideas, so the track record does not count the same profit twice.
  // A leg is part of its position, not a trade of its own, so its Return is what it added to the
  // position's return (its profit over the position's cost base, as a percentage like every
  // other row), never a percentage of its own premium. A leg that leans against the stance is tagged HEDGE.
  const closedLegs = D.ideas.filter((i) => i.position && i.status !== 'closed').flatMap((i) =>
    i.position.legs.filter((l) => l.state !== 'open').map((l) => ({
      ticker: i.ticker, name: i.name, sector: i.sector, status: 'closed', placeholder: i.placeholder,
      stance: l.side === 'Sold' ? 'SHORT' : 'LONG',
      side: l.side.toUpperCase(), expired: l.state === 'expired',
      hedge: (l.side === 'Sold' ? -1 : 1) * (l.kind === 'Put' ? -1 : 1) * (i.stance === 'SHORT' ? -1 : 1) < 0,
      pnl: l.pnl, contrib: i.position.base && l.pnl != null ? l.pnl / i.position.base : null, whole: i.position.ret,
      // Short form so it fits the Ticker column at any width: 55C is the 55-strike call. The full
      // description, with the expiry, is in the tag's tooltip and on the ticker's own page.
      leg: l.kind === 'Shares' ? 'SHARES' : `${l.strike}${l.kind[0]}`,
      legWhen: l.kind === 'Shares' ? '' : fmtDate(l.expiry).slice(3).toUpperCase(),
      legFull: l.kind === 'Shares' ? `${l.side} shares` : `${l.side} ${l.strike} ${l.kind.toLowerCase()} expiring ${fmtDate(l.expiry)}`,
      initiated: l.opened, entry: l.price, exit: l.now, closedOn: l.closed || null,
    })));
  // Two closed legs of one ticker with the same strike and type are told apart by their expiry month.
  closedLegs.filter((r) => r.legWhen && closedLegs.some((o) => o !== r && o.ticker === r.ticker && o.leg === r.leg))
    .forEach((r) => { r.leg += ` ${r.legWhen}`; });
  const qmPool = D.ideas.concat(closedLegs);
  const qmRet = (i) => (i.leg ? i.contrib : calc(i).ret);

  // ---------- quote monitor ----------
  const QM_TABS = [
    { id: 'Active', f: (i) => i.status === 'active' },
    { id: 'Biotech', f: (i) => i.sector === 'Biotech' && i.status !== 'closed' },
    { id: 'Watchlist', f: (i) => i.status === 'watch' },
    // Finished calls only, so its count matches "closed calls" in the track record.
    { id: 'Closed', closed: true, f: (i) => i.status === 'closed' && !i.leg },
    // Legs closed out of a position that is still open are not calls of their own, so they get a
    // tab of their own, shown only while there are any.
    ...(closedLegs.length ? [{ id: 'Closed legs', closed: true, f: (i) => !!i.leg,
      tip: 'Parts of a position that is still open, closed or expired. They are not finished calls, so the track record does not count them.' }] : []),
  ];
  const qmCols = (closed) => [
    { k: 'ticker', label: 'Ticker', left: true, get: (i) => i.ticker },
    { k: 'stance', label: 'Stance', left: true, get: (i) => i.stance },
    { k: 'initiated', label: 'Initiated', get: (i) => i.initiated },
    { k: 'entry', label: 'Entry', get: qmEntry },
    { k: 'target', label: 'Target', get: qmTarget },
    { k: 'mark', label: closed ? 'Exit' : 'Last', get: qmMark },
    { k: 'ret', label: 'Return', get: qmRet },
    // A closed call has no distance to target, so that tab shows the closing date instead.
    closed ? { k: 'closed', label: 'Closed', get: (i) => i.closedOn }
      : { k: 'upside', label: 'To target', get: qmUpside },
  ];

  function renderQM() {
    const tab = QM_TABS.find((t) => t.id === state.tab) || QM_TABS[0];
    const closed = !!tab.closed;
    const cols = qmCols(closed);
    const col = cols.find((c) => c.k === state.sort.k) || cols[0];
    const q = state.qmQ.trim().toUpperCase();
    const rows = qmPool.filter(tab.f)
      .filter((i) => !q || i.ticker.includes(q) || i.name.toUpperCase().includes(q))
      .sort((a, b) => {
        const x = col.get(a), y = col.get(b);
        if (x == null) return 1;
        if (y == null) return -1;
        return (x > y ? 1 : x < y ? -1 : 0) * state.sort.dir;
      });
    $('#qmTabs').innerHTML = QM_TABS.map((t) =>
      `<button class="tab${t.id === tab.id ? ' on' : ''}" data-qmtab="${t.id}"${t.tip ? ` title="${t.tip}"` : ''}>${t.id}<span class="n">${qmPool.filter(t.f).length}</span></button>`).join('');
    // The c-<key> classes let style.css drop the less important columns as the window narrows.
    const head = cols.map((c) =>
      `<th class="c-${c.k}${c.left ? ' l' : ''}" data-sort="${c.k}">${c.label}${state.sort.k === c.k ? `<i class="car${state.sort.dir > 0 ? '' : ' dn'}"></i>` : ''}</th>`).join('');
    const body = rows.map((i) => {
      const c = calc(i);
      return `<tr data-go="#/s/${i.ticker}" class="${state.ctx === i.ticker ? 'sel' : ''}" tabindex="0">
        <td class="c-ticker l tk">${i.ticker}${i.placeholder ? '<span class="ph" title="Some of these figures are placeholders, not the real trades">*</span>' : ''}${i.leg ? `<span class="opt" title="${esc(i.legFull)}. A leg of the ${i.ticker} options position that ${i.expired ? 'has expired' : 'was closed'}. Its profit is also counted in ${i.ticker}'s own row.">${esc(i.leg)}</span>` : ''}</td>
        <td class="c-stance l${i.leg ? '' : ` st-${i.stance.toLowerCase()}`}">${i.leg ? i.side : i.stance}${i.hedge ? '<span class="opt" title="Sold or bought to hedge the position, not as a trade of its own">HEDGE</span>' : ''}${i.position ? '<span class="opt" title="Held through options: Entry and Last are what the position cost and is worth, per share">OPT</span>' : ''}</td>
        <td class="c-initiated">${i.initiated ? fmtDate(i.initiated) : '—'}</td>
        <td class="c-entry">${px(qmEntry(i))}</td><td class="c-target">${px(qmTarget(i))}</td>
        <td class="c-mark ${cls(qmRet(i))}">${px(qmMark(i))}</td>
        <td class="c-ret ${cls(qmRet(i))}"${i.leg ? ` title="This leg added ${signed(i.pnl)} per share to ${i.ticker}'s return of ${pct(i.whole)}. The percentage is measured on the whole position's cost, not on this leg's own premium."` : ''}>${pct(qmRet(i))}</td>
        ${closed ? `<td class="c-closed">${i.closedOn ? fmtDate(i.closedOn) : '—'}</td>` : `<td class="c-upside ${cls(qmUpside(i))}">${pct(qmUpside(i))}</td>`}</tr>`;
    }).join('');
    const none = q ? `No ticker matches “${esc(state.qmQ.trim())}” in this tab` : 'Nothing here yet';
    $('#qmTable').innerHTML = `<thead><tr>${head}</tr></thead><tbody>${body || `<tr><td colspan="8" class="empty">${none}</td></tr>`}</tbody>`;
    $('#qmFoot').textContent = `Prices as of ${fmtDate(D.asOf)}${D.hasSamples ? ' · Includes sample data' : ''}${rows.some((i) => i.placeholder) ? ' · * Placeholder figures' : ''}`;
    tabsMore();
  }
  // The tab strip has no scrollbar. Its right edge fades while tabs are hidden
  // beyond it, and the mouse wheel scrolls it sideways.
  const qmTabs = $('#qmTabs');
  function tabsMore() {
    qmTabs.classList.toggle('more', qmTabs.scrollLeft + qmTabs.clientWidth < qmTabs.scrollWidth - 1);
  }
  qmTabs.addEventListener('scroll', tabsMore);
  qmTabs.addEventListener('wheel', (e) => {
    if (e.deltaX || qmTabs.scrollWidth <= qmTabs.clientWidth) return;
    e.preventDefault();
    qmTabs.scrollLeft += e.deltaY;
  }, { passive: false });
  new ResizeObserver(tabsMore).observe(qmTabs);

  // ---------- updates wire ----------
  function renderWire() {
    const f = state.wire, q = f.q.trim().toLowerCase();
    const list = posts.filter((p) =>
      (f.type === 'ALL' || p.type === f.type) &&
      (!f.ticker || p.tickers.includes(f.ticker)) &&
      (!q || (p.headline + ' ' + p.tickers.join(' ')).toLowerCase().includes(q)));
    $('#wireBody').innerHTML = list.map((p) =>
      `<tr data-go="#/n/${p.id}" class="${p === posts[0] ? 'new' : ''}" tabindex="0">
        <td class="l">${fmtDate(p.date)} <span class="tm">${p.time}</span></td>
        <td class="l"><span class="type t-${p.type}">${p.type}</span></td>
        <td class="l hl" title="${esc(p.headline)}">${esc(p.headline)}</td>
        <td class="l">${p.tickers.map((t) => `<span class="tkchip">${t}</span>`).join('')}</td></tr>`).join('')
      || '<tr><td colspan="4" class="empty">No updates match</td></tr>';
    $('#wireCount').textContent = `Showing ${list.length} of ${posts.length}`;
    $('#wireSub').textContent = f.ticker || 'ALL';
    const chip = $('#wireTk');
    chip.hidden = !f.ticker;
    chip.textContent = f.ticker ? `${f.ticker} ×` : '';
  }
  function clearWire() {
    state.wire = { q: '', type: 'ALL', ticker: null };
    $('#wireQ').value = '';
    $('#wireType').value = 'ALL';
    renderWire();
  }

  // ---------- price chart ----------
  // Draws the real daily closes in i.history ([date, close] pairs written by
  // build_site.py). A ticker without history gets an illustrative path through
  // entry and last instead, and the caption says so.
  function rng(seed) {
    let s = 7;
    for (const ch of seed) s = (s * 31 + ch.charCodeAt(0)) >>> 0;
    return () => (s = (s * 1664525 + 1013904223) >>> 0) / 4294967296;
  }
  function bridge(r, a, b, n, vol) {
    const w = [0];
    for (let k = 1; k < n; k++) w.push(w[k - 1] + (r() - 0.5) * vol);
    return w.map((x, k) => { const t = k / (n - 1); return (a + (b - a) * t) * (1 + x - t * w[n - 1]); });
  }
  const RANGES = [['1M', 1], ['3M', 3], ['6M', 6], ['1Y', 12], ['ALL', 0]];

  // The frame only: a readout line, range buttons, an empty plot and a caption.
  // drawChart fills the plot once the frame is on the page and its width is known.
  function chart(i) {
    const hist = i.history && i.history.length > 1 ? i.history : null;
    return `<div class="chartbox" data-chart="${i.ticker}">
      <div class="chart-bar"><span class="readout"></span><span class="sp"></span>${hist
        ? `<span class="rng">${RANGES.map(([n]) => `<button data-act="range" data-v="${n}" title="Show ${n === 'ALL' ? 'all history' : 'the last ' + n}">${n}</button>`).join('')}</span>` : ''}</div>
      <div class="plot"></div>
      <div class="note">${hist ? `Daily closing prices, ${fmtDate(hist[0][0])} to ${fmtDate(hist[hist.length - 1][0])}. Source: Yahoo Finance.` : 'Illustrative price path, not real price history.'}${hist && (i.checkpoints || []).some(cpSettled) ? ' ✓ and ✗ on the date axis mark settled checkpoints; point at one to read it.' : ''}</div>
    </div>`;
  }

  // The series to plot for the chosen range, and where the call sits in it.
  function chartData(i) {
    const c = calc(i);
    let hist = i.history && i.history.length > 1 ? i.history : null;
    if (hist) {
      const months = (RANGES.find((x) => x[0] === state.range) || RANGES[RANGES.length - 1])[1];
      if (months) {
        const d = new Date(hist[hist.length - 1][0] + 'T00:00:00Z');
        d.setUTCMonth(d.getUTCMonth() - months);
        const cut = hist.filter((p) => p[0] >= d.toISOString().slice(0, 10));
        if (cut.length > 1) hist = cut;
      }
      const called = i.initiated && i.entry != null;
      // The call is marked on the first trading day on or after the initiation date.
      // If that day is before the range shown, only the entry-price line is drawn.
      const before = called && i.initiated < hist[0][0];
      const k = called && !before ? hist.findIndex((p) => p[0] >= i.initiated) : -1;
      return { c, dates: hist.map((p) => p[0]), s: hist.map((p) => p[1]), init: k < 0 ? null : k, entryFrom: before ? 0 : k < 0 ? null : k };
    }
    const r = rng(i.ticker);
    if (i.entry != null) {
      const pre = bridge(r, i.entry * (0.82 + r() * 0.3), i.entry, 30, 0.05);
      return { c, dates: null, s: pre.concat(bridge(r, i.entry, c.mark, 90, 0.05).slice(1)), init: 29, entryFrom: 29 };
    }
    return { c, dates: null, s: bridge(r, i.last * (0.8 + r() * 0.3), i.last, 119, 0.05), init: null, entryFrom: null };
  }

  // A round spacing (1, 2, 2.5, 5 or 10 times a power of ten) for about n axis labels.
  function niceStep(span, n) {
    const raw = span / n, p = 10 ** Math.floor(Math.log10(raw)), f = raw / p;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p;
  }

  // Date labels along the bottom: month starts, thinned to quarters or years when
  // they would crowd; a range too short for that gets evenly spaced days instead.
  function dateTicks(dates, x) {
    const n = dates.length, starts = [];
    dates.forEach((d, k) => { if (k && d.slice(0, 7) !== dates[k - 1].slice(0, 7)) starts.push(k); });
    const perMonth = starts.length > 1 ? (x(starts[starts.length - 1]) - x(starts[0])) / (starts.length - 1) : 0;
    const stride = [1, 2, 3, 6, 12].find((m) => perMonth * m >= 62) || 12;
    const ticks = starts.filter((k) => (+dates[k].slice(5, 7) - 1) % stride === 0).map((k, j) => {
      const mo = +dates[k].slice(5, 7) - 1, yy = dates[k].slice(2, 4);
      return { k, label: stride === 12 ? dates[k].slice(0, 4) : MONTHS[mo] + (mo === 0 || j === 0 ? ` '${yy}` : '') };
    });
    if (ticks.length >= 3) return ticks;
    const m = Math.min(5, n - 1);
    return Array.from({ length: m + 1 }, (_, j) => {
      const k = Math.round((j * (n - 1)) / m);
      return { k, label: `${+dates[k].slice(8, 10)} ${MONTHS[+dates[k].slice(5, 7) - 1]}` };
    });
  }

  function drawChart(box) {
    const i = ideaBy[box.dataset.chart], plot = $('.plot', box);
    if (!i || !plot) return;
    const m = chartData(i), s = m.s, c = m.c;
    // Drawn at the plot's real pixel width, so text stays one size however wide the window is.
    const W = Math.max(300, Math.round(plot.clientWidth)), H = W < 480 ? 220 : 280;
    const L = 10, R = 64, T = 24, B = 24, xe = W - R, yb = H - B;
    const showTarget = i.target != null && i.status === 'active';
    // The scale covers the prices, the target line, the entry price and the last-price tag.
    const marks = [c.mark].concat(showTarget ? [i.target] : [], m.entryFrom != null ? [i.entry] : []).filter((v) => v != null);
    let lo = Math.min(...s, ...marks), hi = Math.max(...s, ...marks);
    const pad = (hi - lo || hi || 1) * 0.08;
    lo = Math.max(0, lo - pad); hi += pad;
    const x = (k) => L + (k / (s.length - 1)) * (xe - L);
    const y = (v) => T + ((hi - v) / (hi - lo)) * (yb - T);
    const f = (v) => v.toFixed(1);
    const yl = y(c.mark), yt = showTarget ? y(i.target) : null;
    const gid = `cg-${i.ticker}`;

    // Price axis: gridlines with labels in the right-hand gutter.
    let grid = '';
    const step = niceStep(hi - lo, H < 250 ? 4 : 5);
    for (let v = Math.ceil(lo / step) * step; v <= hi; v += step) {
      const yy = y(v);
      grid += `<line x1="${L}" x2="${xe + 4}" y1="${f(yy)}" y2="${f(yy)}" stroke="#1f1f1f"/>`;
      if (Math.abs(yy - yl) > 13 && (yt == null || Math.abs(yy - yt) > 13)) grid += `<text x="${xe + 8}" y="${f(yy + 4)}" fill="#8f8f8f" font-size="11">${px(v)}</text>`;
    }
    // Date axis.
    if (m.dates) {
      dateTicks(m.dates, x).forEach(({ k, label }) => {
        const xx = x(k), anchor = xx < L + 26 ? 'start' : xx > xe - 26 ? 'end' : 'middle';
        grid += `<line x1="${f(xx)}" x2="${f(xx)}" y1="${T}" y2="${yb + 4}" stroke="#1f1f1f"/>
          <text x="${f(xx)}" y="${yb + 16}" fill="#8f8f8f" font-size="11" text-anchor="${anchor}">${label}</text>`;
      });
    }

    const pts = s.map((v, k) => `${f(x(k))},${f(y(v))}`).join(' ');
    let out = `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Price chart for ${i.ticker}">
      <defs><linearGradient id="${gid}" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0" stop-color="#45a873" stop-opacity=".28"/><stop offset="1" stop-color="#45a873" stop-opacity="0"/></linearGradient></defs>
      ${grid}
      <polygon points="${L},${yb} ${pts} ${xe},${yb}" fill="url(#${gid})"/>
      <polyline points="${pts}" fill="none" stroke="#45a873" stroke-width="1.5" stroke-linejoin="round"/>
      <line x1="${L}" x2="${xe}" y1="${yb}" y2="${yb}" stroke="#474747"/>
      <line x1="${xe}" x2="${xe}" y1="${T}" y2="${yb}" stroke="#474747"/>`;
    if (m.entryFrom != null) {
      out += `<line x1="${f(x(m.entryFrom))}" x2="${xe}" y1="${f(y(i.entry))}" y2="${f(y(i.entry))}" stroke="#8f8f8f" stroke-dasharray="2 3"/>`;
    }
    if (m.init != null) {
      const xi = x(m.init), flip = xi + 190 > W;
      out += `<line x1="${f(xi)}" x2="${f(xi)}" y1="${T - 6}" y2="${yb}" stroke="#8f8f8f" stroke-dasharray="3 3"/>
        <circle cx="${f(xi)}" cy="${f(y(i.entry))}" r="3.5" fill="#000" stroke="#d6d6d6" stroke-width="1.5"/>
        <text x="${f(flip ? xi - 5 : xi + 5)}" y="${T - 9}" fill="#d6d6d6" font-size="11"${flip ? ' text-anchor="end"' : ''}>${i.stance === 'SHORT' ? 'SHORT' : 'LONG'} ${fmtDate(i.initiated)} @ ${px(i.entry)}</text>`;
    }
    // Checkpoints settled inside the range shown: a mark on the date axis; the readout names them.
    const cpAt = {};
    if (m.dates && i.checkpoints) {
      const first = m.dates[0], last = m.dates[m.dates.length - 1];
      i.checkpoints.filter((c) => cpSettled(c) && c.settled >= first && c.settled <= last).forEach((c) => {
        const k = m.dates.findIndex((d) => d >= c.settled);
        (cpAt[k] = cpAt[k] || []).push(c);
      });
      Object.entries(cpAt).forEach(([k, cs]) => {
        const xx = x(+k), good = cs.every(cpGood);
        out += `<line x1="${f(xx)}" x2="${f(xx)}" y1="${yb - 6}" y2="${yb}" stroke="#d6d6d6"/>
          <text x="${f(xx)}" y="${yb - 9}" fill="${good ? '#45a873' : '#ff4b5c'}" font-size="10" text-anchor="middle">${good ? '✓' : '✗'}</text>`;
      });
    }
    if (showTarget) {
      // The target tag steps aside when it would sit on the last-price tag.
      const tag = Math.abs(yt - yl) < 17 ? yl + (yt <= yl ? -17 : 17) : yt;
      out += `<line x1="${L}" x2="${xe}" y1="${f(yt)}" y2="${f(yt)}" stroke="#ffb454" stroke-dasharray="4 4"/>
        <text x="${xe - 5}" y="${f(yt - 5)}" fill="#ffb454" font-size="10" text-anchor="end" letter-spacing=".08em">TARGET</text>
        <rect x="${xe + 1.5}" y="${f(tag - 8)}" width="${R - 8}" height="16" fill="#000" stroke="#ffb454"/>
        <text x="${xe + 8}" y="${f(tag + 4)}" fill="#ffb454" font-size="11">${px(i.target)}</text>`;
    }
    out += `<rect x="${xe + 1}" y="${f(yl - 8)}" width="${R - 7}" height="16" fill="#45a873"/>
      <text x="${xe + 8}" y="${f(yl + 4)}" fill="#03140b" font-size="11" font-weight="700">${px(c.mark)}</text>
      <g class="xh" visibility="hidden" pointer-events="none">
        <g class="xh-x"><line y1="${T}" y2="${yb}" stroke="#8f8f8f" stroke-dasharray="2 2"/></g>
        <g class="xh-y"><line x1="${L}" x2="${xe}" stroke="#8f8f8f" stroke-dasharray="2 2"/>
          <rect x="${xe + 1}" y="-8" width="${R - 7}" height="16" fill="#474747"/><text class="xh-px" x="${xe + 8}" y="4" fill="#fff" font-size="11"></text></g>
        <circle class="xh-dot" r="3.5" fill="#000" stroke="#45a873" stroke-width="1.5"/>
        <g class="xh-d"${m.dates ? '' : ' display="none"'}><rect x="-36" y="${yb + 4}" width="72" height="16" fill="#474747"/>
          <text class="xh-date" y="${yb + 16}" fill="#fff" font-size="11" text-anchor="middle"></text></g>
      </g></svg>`;
    plot.innerHTML = out;
    box._m = { i, s, dates: m.dates, W, L, xe, x, y, cpAt };
    box.querySelectorAll('.rng button').forEach((b) => b.classList.toggle('on', b.dataset.v === state.range));
    chartReadout(box, s.length - 1);
    if (!box._wired) {
      box._wired = true;
      plot.addEventListener('pointermove', (e) => chartHover(box, e));
      plot.addEventListener('pointerleave', () => {
        const xh = $('.xh', box);
        if (xh) xh.setAttribute('visibility', 'hidden');
        chartReadout(box, box._m.s.length - 1);
      });
    }
  }
  const drawCharts = () => document.querySelectorAll('.chartbox').forEach(drawChart);

  // The line above the plot: the date and close under the pointer, or the latest when idle.
  function chartReadout(box, k) {
    const m = box._m, i = m.i, v = m.s[k];
    let html = `<b>${i.ticker}</b>`;
    if (m.dates) {
      html += ` ${fmtDate(m.dates[k])} · Close <b>${px(v)}</b>`;
      if (k > 0 && m.s[k - 1]) { const d = v / m.s[k - 1] - 1; html += ` <span class="${cls(d)}">${pct(d)}</span>`; }
      if (i.entry != null && i.initiated && m.dates[k] >= i.initiated) {
        const r = (i.stance === 'SHORT' ? i.entry - v : v - i.entry) / i.entry;
        html += ` · ${i.position ? 'Stock since call' : 'Since call'} <span class="${cls(r)}">${pct(r)}</span>`;
      }
      if (m.cpAt[k]) html += ' · ' + m.cpAt[k].map((c) => `${cpMark(c)} <b>${esc(c.what)}</b>`).join(' · ');
    } else html += ' · Illustrative path';
    $('.readout', box).innerHTML = html;
  }

  // Crosshair: snaps to the nearest day and tags its price and date on the axes.
  function chartHover(box, e) {
    const m = box._m, svg = $('svg', box);
    if (!m || !svg) return;
    const rc = svg.getBoundingClientRect(), n = m.s.length - 1;
    const sx = (e.clientX - rc.left) * (m.W / rc.width);
    const k = Math.max(0, Math.min(n, Math.round(((sx - m.L) / (m.xe - m.L)) * n)));
    const X = m.x(k).toFixed(1), Y = m.y(m.s[k]).toFixed(1);
    $('.xh', svg).setAttribute('visibility', 'visible');
    $('.xh-x', svg).setAttribute('transform', `translate(${X},0)`);
    $('.xh-y', svg).setAttribute('transform', `translate(0,${Y})`);
    $('.xh-dot', svg).setAttribute('transform', `translate(${X},${Y})`);
    $('.xh-px', svg).textContent = px(m.s[k]);
    if (m.dates) {
      $('.xh-d', svg).setAttribute('transform', `translate(${Math.max(m.L + 36, Math.min(m.xe - 36, +X))},0)`);
      $('.xh-date', svg).textContent = fmtDate(m.dates[k]);
    }
    chartReadout(box, k);
  }

  // ---------- security pages ----------
  const list = (xs) => `<ul class="tight">${xs.map((x) => `<li>${esc(x)}</li>`).join('')}</ul>`;
  const kv = (rows) => `<table class="kv">${rows.map(([k, v, c]) => `<tr><td>${esc(k)}</td><td class="${c || ''}">${esc(v)}</td></tr>`).join('')}</table>`;
  const tickerPosts = (t) => posts.filter((p) => p.tickers.includes(t));

  // ---------- checkpoints ----------
  // i.checkpoints come from the Checkpoints section of the ticker note (build_site.py reads
  // them): NEED is what the call needs to happen, EXIT is what would make me close it. A
  // settled checkpoint went the call's way when a NEED was met or an EXIT did not happen.
  const cpSettled = (c) => c.status === 'met' || c.status === 'missed';
  const cpGood = (c) => (c.status === 'met') === (c.kind === 'NEED');
  const cpMark = (c) => (cpSettled(c) ? (cpGood(c) ? '<span class="up">✓</span>' : '<span class="down">✗</span>') : c.status === 'dropped' ? '–' : '○');
  const anyCheckpoints = D.ideas.some((i) => i.checkpoints);
  // A checkpoint written down after its answer was known (hindsight) is shown but not counted.
  function cpTally(cps) {
    const settled = cps.filter((c) => cpSettled(c) && !c.hindsight), good = settled.filter(cpGood).length;
    return { settled: settled.length, good, bad: settled.length - good, open: cps.filter((c) => c.status === 'pending').length };
  }
  function cpOutcome(c) {
    if (c.status === 'pending') return c.overdue ? 'awaiting update' : c.auto ? 'watching the price' : 'open';
    if (c.status === 'dropped') return 'dropped' + (c.settled ? ' ' + fmtDate(c.settled) : '');
    return `${c.kind === 'NEED' ? c.status : c.status === 'met' ? 'happened' : 'did not happen'} ${fmtDate(c.settled)}`;
  }
  // One row: the mark, what has to happen with its dates beneath, and the outcome
  // (a link to the update about it, when there is one). tk adds a ticker column; withKind
  // starts the dates line with NEED or EXIT, for lists not split into the two groups.
  function cpRow(c, tk, withKind) {
    const out = esc(cpOutcome(c));
    const detail = (withKind ? [c.kind] : []).concat([c.due === 'any time' ? 'any time' : `by ${c.due}`, `set ${fmtDate(c.set)}`])
      .concat(c.price != null ? [`closed at ${px(c.price)}`] : [], c.hindsight ? ['added after it was settled, so not counted'] : []).join(' · ');
    return `<tr class="${c.overdue ? 'late' : ''}"><td class="cp-m">${cpMark(c)}</td>${tk ? `<td class="cp-t"><button class="tkchip" data-go="#/s/${tk}">${tk}</button></td>` : ''}
      <td>${esc(c.what)}<div class="d">${esc(detail)}</div></td>
      <td class="cp-o">${c.post ? `<button data-go="#/n/${c.post}">${out} ›</button>` : out}</td></tr>`;
  }
  const cpKey = '<p class="note">Each checkpoint is dated from the day it first appeared on this site. ✓ went the call’s way, ✗ went against it.</p>';
  function checkpointsBlock(i) {
    const cps = i.checkpoints;
    if (!cps) return '';
    const t = cpTally(cps);
    const hit = i.status !== 'closed' && cps.find((c) => c.kind === 'EXIT' && c.status === 'met');
    const group = (kind, title) => {
      const rows = cps.filter((c) => c.kind === kind);
      return rows.length ? `<tr class="cp-sec"><td colspan="3">${title}</td></tr>${rows.map((c) => cpRow(c)).join('')}` : '';
    };
    return `<h2 class="cp-h">Checkpoints<span>${t.good} for the call · ${t.bad} against · ${t.open} open</span></h2>
      ${hit ? `<p class="cp-flag">An exit checkpoint was met on ${fmtDate(hit.settled)} and the call is still open.</p>` : ''}
      <table class="mini cps">${group('NEED', 'What the call needs')}${group('EXIT', 'What would make me close it')}</table>${cpKey}`;
  }
  // The CP page: the open checkpoints, soonest first, then the settled and dropped ones,
  // latest first. Across every call (the CP command), or for one ticker (ABVX CP).
  function viewCheckpoints(tk) {
    const all = (tk ? [ideaBy[tk]] : D.ideas).flatMap((i) => (i.checkpoints || []).map((c) => ({ c, i })));
    // Across every call, a closed call's open checkpoints are left out; its own page shows them.
    const open = all.filter(({ c, i }) => c.status === 'pending' && (tk || i.status !== 'closed'))
      .sort((a, b) => (a.c.by || '9999').localeCompare(b.c.by || '9999') || a.i.ticker.localeCompare(b.i.ticker));
    const done = all.filter(({ c }) => c.status !== 'pending')
      .sort((a, b) => (b.c.settled || b.c.set).localeCompare(a.c.settled || a.c.set));
    const t = cpTally(all.map(({ c }) => c));
    const table = (rows) => `<table class="mini cps">${rows.map(({ c, i }) => cpRow(c, tk ? null : i.ticker, true)).join('')}</table>`;
    return `<div class="page"><h1>${tk ? `${tk} checkpoints` : 'Checkpoints'}</h1>
      <p class="lead">What ${tk ? 'the call' : 'each call'} needs to happen (NEED), and what would make me close it (EXIT), written down before the answer is known.</p>
      ${t.settled ? `<p class="muted">${t.good} of ${t.settled} settled checkpoints went the call’s way.</p>` : ''}
      ${all.length ? `<h2>Open, soonest first</h2>${open.length ? table(open) : '<p class="muted">None open.</p>'}` : '<p class="muted">No checkpoints for this name yet.</p>'}
      ${done.length ? `<h2>Settled, latest first</h2>${table(done)}` : ''}
      ${cpKey}</div>`;
  }

  function secNav(i, on) {
    const tab = (id, label, sub) => `<button class="tab${on === id ? ' on' : ''}" data-go="#/s/${i.ticker}${sub}">${label}</button>`;
    return `<div class="secnav">${tab('des', 'DES', '')}${tab('thesis', 'THESIS', '/thesis')}${tab('model', 'MODEL', '/model')}${i.checkpoints ? tab('cp', 'CP', '/cp') : ''}<button class="tab" data-go="#/s/${i.ticker}/n">N</button></div>`;
  }

  function viewDes(i) {
    const c = calc(i), tp = tickerPosts(i.ticker), pos = i.position;
    const call = [['Stance', i.stance], ['Conviction', i.conviction || '—'], ['Initiated', i.initiated ? fmtDate(i.initiated) : '—']].concat(pos
      // Held through options: the position's own cost, value and return, then the stock's prices for reference.
      ? [['Held through', 'Options'], [pos.cost < 0 ? 'Net received' : 'Net cost', px(Math.abs(pos.cost))], ['Value now', px(pos.value)],
        ['Return', pct(c.ret), cls(c.ret)], ['Stock then', px(i.entry)], [i.status === 'closed' ? 'Stock at close' : 'Stock now', px(c.mark)]]
        .concat(i.target != null ? [['Stock target', px(i.target)]] : [])
      : [['Entry', px(i.entry)], ['Target', px(i.target)], [i.status === 'closed' ? 'Exit' : 'Last', px(c.mark)],
        ['Return', pct(c.ret), cls(c.ret)]]);
    if (i.status === 'active' && !pos) call.push(['To target', pct(c.upside), cls(c.upside)]);
    if (i.status === 'closed') call.push(['Closed', fmtDate(i.closedOn)]);
    if (i.catalyst) call.push(['Next catalyst', i.catalyst.when]);
    if (i.checkpoints) {
      const t = cpTally(i.checkpoints);
      call.push(['Checkpoints', t.settled ? `${t.good}/${t.settled} for the call` : `${t.open} open`]);
    }
    return `<div class="des">
      <div>
        <div class="sec-top">
          <div>
            <div class="sec-id"><span class="tkbig">${i.ticker}</span><span class="badge ${i.stance.toLowerCase()}">${i.stance}</span>${i.status === 'closed' ? '<span class="badge closed">CLOSED</span>' : ''}${i.placeholder ? '<span class="badge closed" title="Some of these figures are placeholders, not the real trades">PLACEHOLDER FIGURES</span>' : ''}</div>
            <div class="sec-name">${esc(i.name)}</div>
            <div class="muted">${esc(i.sector)} · ${esc(i.blurb)}</div>
          </div>
          <div class="sec-px"><span class="px">${px(c.mark)}</span><span class="${cls(c.ret)}">${pos ? (c.ret == null ? 'options position' : pct(c.ret) + ' on the position') : c.ret == null ? 'no call yet' : pct(c.ret) + ' since call'}</span></div>
        </div>
        <div class="jump">
          <button class="btn" data-go="#/s/${i.ticker}/thesis">THESIS ›</button>
          <button class="btn" data-go="#/s/${i.ticker}/model">MODEL ›</button>
          <button class="btn" data-go="#/s/${i.ticker}/n">N ›</button>
        </div>
        <h2>Thesis in brief</h2>${list(i.summary)}
        ${checkpointsBlock(i)}
        ${pos ? positionBlock(i) : ''}
        <h2>${pos ? 'Share price since the call' : 'Price since the call'}</h2>${chart(i)}
        ${i.drivers ? `<div class="cols2"><div><h2>Key drivers</h2>${list(i.drivers)}</div><div><h2>Risks</h2>${list(i.risks)}</div></div>` : ''}
        ${i.catalysts ? `<h2>Catalysts</h2><table class="mini">${i.catalysts.map(([w, e]) => `<tr><td>${esc(w)}</td><td>${esc(e)}</td></tr>`).join('')}</table>` : ''}
        <h2>Updates on ${i.ticker}</h2>
        <table class="mini posts">${tp.map((p) => `<tr><td>${fmtDate(p.date)}</td><td><span class="type t-${p.type}">${p.type}</span></td><td><button data-go="#/n/${p.id}">${esc(p.headline)}</button></td></tr>`).join('') || '<tr><td>None yet</td></tr>'}</table>
      </div>
      <aside class="des-side">
        <h2>My call</h2>${kv(call)}
        <h2>Company</h2>${kv(i.stats || [])}
        ${i.scenarios ? `<h2>Valuation</h2>${kv(i.scenarios.map((s) => [`${s.name} (${Math.round(s.prob * 100)}%)`, px(s.value)]).concat([['Weighted', px(i.target)]]))}` : ''}
      </aside>
    </div>`;
  }

  // ---------- options position ----------
  // i.position is worked out by build_site.py from the legs in the ticker note: each leg's
  // price now, the position's cost, value and return, and its profit or loss at expiry for a
  // range of share prices (p.payoff). Every figure is per share, for one unit of the position.
  const signed = (v) => (v == null ? '—' : (v > 0 ? '+' : v < 0 ? '−' : '') + Math.abs(v).toFixed(2));
  const isoToDate = (s) => String(s).replace(/\d{4}-\d{2}-\d{2}/g, (d) => fmtDate(d));

  function positionBlock(i) {
    const p = i.position;
    const legs = p.legs.map((l) => `<tr class="${l.state === 'open' ? '' : 'off'}">
      <td class="l">${l.side} ${l.qty !== 1 ? `${l.qty} × ` : ''}${l.kind.toLowerCase()}</td>
      <td>${l.strike == null ? '—' : px(l.strike)}</td><td>${l.expiry ? fmtDate(l.expiry) : '—'}</td>
      <td>${px(l.price)}</td><td>${px(l.now)}</td><td class="${cls(l.pnl)}">${signed(l.pnl)}</td>
      <td class="l how">${esc(isoToDate(l.how))}</td></tr>`).join('');
    const left = [[p.cost < 0 ? 'Net received' : 'Net cost', px(Math.abs(p.cost))], ['Value now', px(p.value)],
      ['Profit or loss', signed(p.pnl), cls(p.pnl)], [p.basis ? `Return on ${p.basis}` : 'Return', pct(p.ret), cls(p.ret)]];
    const right = p.payoff ? [['Best case at expiry', p.maxGain == null ? 'Unlimited' : signed(p.maxGain)],
      ['Worst case at expiry', p.maxLoss == null ? 'Unlimited' : signed(-p.maxLoss)],
      ['Breakeven at expiry', p.breakevens.length ? p.breakevens.map(px).join(', ') : '—'],
      [i.status === 'closed' ? 'Stock at close' : 'Stock now', px(p.stock)]] : null;
    return `<h2>Position</h2>
      <div class="scroll-x"><table class="grid pos"><thead><tr><th class="l">Leg</th><th>Strike</th><th>Expiry</th><th>Price</th><th>Now</th><th>P/L</th><th class="l">Priced by</th></tr></thead>
        <tbody>${legs}</tbody></table></div>
      <div class="${right ? 'cols2 ' : ''}pos-sum"><div>${kv(left)}</div>${right ? `<div>${kv(right)}</div>` : ''}</div>
      ${p.payoff ? payoffChart(i) : ''}
      <p class="note pos-note">${i.placeholder ? 'Placeholder figures: the real trade details are still to come. ' : ''}${esc(isoToDate(p.valued))} Figures are per share, for one unit of the position, as of ${fmtDate(p.asOf)}.</p>`;
  }

  // Profit or loss at expiry against the share price: straight lines between the strikes.
  function payoffChart(i) {
    const p = i.position, pts = p.payoff;
    const W = 640, H = 210, L = 10, R = 58, T = 20, B = 26, xe = W - R, yb = H - B;
    const x0 = pts[0][0], x1 = pts[pts.length - 1][0], vs = pts.map((q) => q[1]);
    let lo = Math.min(0, ...vs), hi = Math.max(0, ...vs);
    const pad = (hi - lo || 1) * 0.12;
    lo -= pad; hi += pad;
    const x = (s) => L + ((s - x0) / (x1 - x0)) * (xe - L), y = (v) => T + ((hi - v) / (hi - lo)) * (yb - T);
    const f = (v) => v.toFixed(1), y0 = y(0), id = `po-${i.ticker}`;
    const line = pts.map(([s, v]) => `${f(x(s))},${f(y(v))}`).join(' ');
    const area = `${L},${f(y0)} ${line} ${xe},${f(y0)}`;
    // Strikes along the bottom; a label that would touch the one before it is dropped.
    let lastX = -99, axis = '';
    pts.slice(1, -1).forEach(([s]) => {
      const xx = x(s);
      axis += `<line x1="${f(xx)}" x2="${f(xx)}" y1="${T}" y2="${yb + 4}" stroke="#1f1f1f"/>`;
      if (xx - lastX > 40) { axis += `<text x="${f(xx)}" y="${yb + 16}" fill="#8f8f8f" font-size="11" text-anchor="middle">${px(s)}</text>`; lastX = xx; }
    });
    // Profit or loss scale on the right: the best and worst points drawn, and zero.
    const top = Math.max(...vs), bot = Math.min(...vs);
    const tags = [[0, '0.00']].concat(top > 0 ? [[top, signed(top)]] : [], bot < 0 ? [[bot, signed(bot)]] : [])
      .filter(([v], n, all) => n === 0 || Math.abs(y(v) - y0) > 13)
      .map(([v, t]) => `<text x="${xe + 8}" y="${f(y(v) + 4)}" fill="#8f8f8f" font-size="11">${t}</text>`).join('');
    let marks = '';
    p.breakevens.filter((s) => s > x0 && s < x1).forEach((s) => {
      marks += `<circle cx="${f(x(s))}" cy="${f(y0)}" r="3" fill="#000" stroke="#d6d6d6" stroke-width="1.5"><title>Breakeven ${px(s)}</title></circle>`;
    });
    if (p.stock > x0 && p.stock < x1) {
      const xs = x(p.stock), flip = xs + 110 > xe;
      marks += `<line x1="${f(xs)}" x2="${f(xs)}" y1="${T - 6}" y2="${yb}" stroke="#8f8f8f" stroke-dasharray="3 3"/>
        <text x="${f(flip ? xs - 5 : xs + 5)}" y="${T - 8}" fill="#d6d6d6" font-size="11"${flip ? ' text-anchor="end"' : ''}>${i.status === 'closed' ? 'Stock at close' : 'Stock now'} ${px(p.stock)}</text>`;
    }
    return `<div class="payoff-h"><b>Profit or loss at expiry</b><span class="muted">per share, by share price on the expiry day</span></div>
      <svg class="chart payoff" viewBox="0 0 ${W} ${H}" role="img" aria-label="Profit or loss at expiry for the ${i.ticker} position">
        <defs><clipPath id="${id}-up"><rect x="${L}" y="${T}" width="${xe - L}" height="${f(Math.max(0, y0 - T))}"/></clipPath>
          <clipPath id="${id}-dn"><rect x="${L}" y="${f(y0)}" width="${xe - L}" height="${f(Math.max(0, yb - y0))}"/></clipPath></defs>
        ${axis}
        <polygon points="${area}" fill="#45a873" fill-opacity=".22" clip-path="url(#${id}-up)"/>
        <polygon points="${area}" fill="#ff4b5c" fill-opacity=".2" clip-path="url(#${id}-dn)"/>
        <line x1="${L}" x2="${xe}" y1="${f(y0)}" y2="${f(y0)}" stroke="#474747"/>
        <line x1="${xe}" x2="${xe}" y1="${T}" y2="${yb}" stroke="#474747"/>
        <polyline points="${line}" fill="none" stroke="#d6d6d6" stroke-width="1.5" stroke-linejoin="round"/>
        ${marks}${tags}</svg>`;
  }

  function viewThesis(i) {
    const c = calc(i);
    // i.thesisDoc is a thesis written as its own Obsidian note. build_site.py turns it
    // into page text (thesis_note.py escapes everything in it), so its html is used as it is.
    const doc = i.thesisDoc;
    const body = doc ? `<div class="md">${doc.html}</div>`
      : i.thesis
      ? i.thesis.map(([h, p]) => `<h2>${esc(h)}</h2><p>${esc(p)}</p>`).join('')
      : `<h2>Summary</h2>${list(i.summary)}<p class="note" style="margin-top:12px">No long-form thesis for this name yet.</p>`;
    return `<article class="reader f${state.font}">
      <div class="reader-bar"><button data-act="copy">Copy link</button><span class="sp"></span>${fontButtons()}<button data-act="print">Export PDF</button></div>
      <div class="meta">${doc && doc.title ? `${i.ticker} · ` : ''}${i.stance}${!i.initiated ? '' : i.position ? ` · OPTIONS · INITIATED ${fmtDate(i.initiated).toUpperCase()}` : ` · INITIATED ${fmtDate(i.initiated).toUpperCase()} @ ${px(i.entry)} · TARGET ${px(i.target)}`}${c.ret != null ? ` · <span class="${cls(c.ret)}">${pct(c.ret)}</span>` : ''}</div>
      <h1>${doc && doc.title ? esc(doc.title) : `${i.ticker}: ${esc(i.name)}`}</h1>
      <p class="muted">${esc(i.blurb)}${doc ? ` · Updated ${fmtDate(doc.updated)}` : ''}</p>
      ${body}
    </article>`;
  }

  function fmtCell(v, kind) {
    if (v == null) return '—';
    if (kind === 'pct') return (v * 100).toFixed(1) + '%';
    if (kind === 'eps') return v < 0 ? `(${Math.abs(v).toFixed(2)})` : v.toFixed(2);
    if (kind === 'plain') return String(v);
    const n = Math.round(Math.abs(v)).toLocaleString('en-US');
    return v < 0 ? `(${n})` : n;
  }
  // The scenario table sits under either kind of model. Empty when the note has no scenarios.
  function scenarioTable(i) {
    if (!i.scenarios || !i.scenarios.length) return '';
    const sc = i.scenarios.map((s) => `<tr><td class="l">${esc(s.name)}</td><td>${Math.round(s.prob * 100)}%</td><td>${px(s.value)}</td><td>${px(s.prob * s.value)}</td><td class="l cat">${esc(s.note)}</td></tr>`).join('');
    return `<h2>Scenario valuation</h2>
      <div class="scroll-x"><table class="grid"><thead><tr><th class="l">Scenario</th><th>Probability</th><th>Value / share</th><th>Weighted</th><th class="l">Assumption</th></tr></thead>
        <tbody>${sc}<tr class="strong"><td class="l">Target</td><td>100%</td><td></td><td>${px(i.target)}</td><td class="l cat">Probability-weighted value; last ${px(i.last)}</td></tr></tbody></table></div>`;
  }
  // A ticker's model is its Excel workbook when it has one (i.workbook), otherwise
  // the table typed in its note (i.model).
  function viewModel(i) {
    if (i.workbook && i.workbook.sheets.length) {
      if (state.wb.ticker !== i.ticker) state.wb = { ticker: i.ticker, sheet: 0, row: wbDefaultRow(i.workbook.sheets[0]) };
      return `<div class="modelv"><div id="wbView">${wbView(i)}</div>${scenarioTable(i)}</div>`;
    }
    if (!i.model) {
      const none = i.status === 'closed' ? 'Model archived when the call was closed.'
        : i.status === 'watch' ? 'No model yet. This name is on the watchlist.' : 'No model yet.';
      return `<div class="modelv"><p class="muted">${none}</p>${scenarioTable(i)}</div>`;
    }
    const m = i.model;
    const head = m.periods.map((p, k) => `<th class="${k >= m.est ? 'est' : ''}">${p}</th>`).join('');
    const rows = m.rows.map((r) => `<tr class="${r.strong ? 'strong' : ''} ${r.kind === 'pct' ? 'pct' : ''}"><td class="l">${esc(r.label)}</td>${
      r.v.map((v, k) => `<td class="${k >= m.est ? 'est' : ''} ${v != null && v < 0 ? 'neg' : ''}">${fmtCell(v, r.kind)}</td>`).join('')}</tr>`).join('');
    return `<div class="modelv">
      <div class="model-bar"><b>${i.ticker} income model</b><span class="muted">${esc(m.unit)} · shaded columns are my estimates</span></div>
      <div class="scroll-x"><table class="grid"><thead><tr><th class="l">Line</th>${head}</tr></thead><tbody>${rows}</tbody></table></div>
      ${scenarioTable(i)}
    </div>`;
  }

  // ---------- Excel model ----------
  // i.workbook is written by build_site.py from the Excel file in the vault's
  // Models folder: { name, updated, file, sheets: [{ name, rows, hdr, truncated }] }.
  // A cell is 0 (empty), "text", or [text, value, flags]; excel_model.py lists the flags.
  const wbText = (c) => (c === 0 ? '' : typeof c === 'string' ? c : c[0]);
  const wbNum = (c) => (Array.isArray(c) && typeof c[1] === 'number' ? c[1] : null);
  const wbChartable = (row) => row.filter((c) => wbNum(c) != null).length >= 2;
  const wbLabel = (row) => wbText(row.find((c) => c !== 0 && wbNum(c) == null) || 0);
  // The row charted when a sheet opens: revenue, if the sheet has such a row. On
  // any other sheet nothing is charted until a row is clicked.
  function wbDefaultRow(sheet) {
    const k = sheet.rows.findIndex((r, n) => wbChartable(r) && n !== sheet.hdr && /revenue|sales/i.test(wbLabel(r)));
    return k < 0 ? null : k;
  }
  function wbRowHtml(sheet, row, k) {
    let html = '';
    for (let j = 0; j < row.length; j++) {
      const c = row[j], num = wbNum(c), flags = Array.isArray(c) && c[2] ? c[2] : '';
      // Text runs over the empty cells to its right, as it does in Excel.
      let span = 1;
      if (c !== 0 && num == null) while (j + span < row.length && row[j + span] === 0) span++;
      // Headings over the number columns line up with the numbers below them.
      const cls = [num == null && !(k === sheet.hdr && j > 0) ? 'l' : '', j === 0 && span === 1 ? 'stick' : '',
        // Plain text to the right of the label column (a source or reasoning note) is shown quieter.
        c !== 0 && num == null && j > 0 && k !== sheet.hdr && !flags.includes('b') ? 'aside' : '',
        flags.includes('b') ? 'b' : '', flags.includes('i') ? 'i' : '', flags.includes('n') ? 'neg' : '',
        flags.includes('B') ? 'cB' : flags.includes('R') ? 'cR' : flags.includes('G') ? 'cG' : ''].filter(Boolean).join(' ');
      const indent = +(flags.match(/[1-6]/) || [0])[0];
      html += `<td class="${cls}"${span > 1 ? ` colspan="${span}"` : ''}${indent ? ` style="padding-left:${6 + indent * 12}px"` : ''}>${esc(wbText(c))}</td>`;
      j += span - 1;
    }
    const pick = wbChartable(row) && k !== sheet.hdr;
    return `<tr class="${k === sheet.hdr ? 'hdr' : ''}${pick ? ' pick' : ''}${pick && k === state.wb.row ? ' sel' : ''}"${pick ? ` data-wbrow="${k}" tabindex="0"` : ''}>${html}</tr>`;
  }
  // One row of a sheet as bars, labelled from the sheet's header row.
  function wbChart(sheet, k) {
    const row = sheet.rows[k], hdr = sheet.hdr != null && sheet.hdr !== k ? sheet.rows[sheet.hdr] : null;
    const pts = [];
    row.forEach((c, j) => {
      const v = wbNum(c);
      if (v != null) pts.push({ v, text: wbText(c).trim(), label: hdr && hdr[j] !== 0 && hdr[j] != null ? wbText(hdr[j]) : '' });
    });
    const W = 640, H = 190, L = 10, R = 10, T = 20, B = 34;
    const lo = Math.min(0, ...pts.map((p) => p.v)), hi = Math.max(0, ...pts.map((p) => p.v)), range = hi - lo || 1;
    const y = (v) => T + ((hi - v) / range) * (H - T - B);
    const slot = (W - L - R) / pts.length, bw = Math.min(46, slot * 0.62);
    // Labels are thinned out when there are too many bars to label each one.
    const every = Math.max(1, Math.ceil(pts.length / Math.floor((W - L - R) / 62)));
    const room = Math.max(4, Math.floor((slot * every - 6) / 6.1)); // characters of heading that fit under a bar
    const bars = pts.map((p, n) => {
      const cx = L + slot * (n + 0.5), top = y(Math.max(p.v, 0)), h = Math.max(1, Math.abs(y(p.v) - y(0)));
      const labels = n % every ? '' : `<text x="${cx.toFixed(1)}" y="${(p.v >= 0 ? top - 4 : top + h + 11).toFixed(1)}" text-anchor="middle" font-size="10" fill="#d6d6d6">${esc(p.text)}</text>
        <text x="${cx.toFixed(1)}" y="${H - 8}" text-anchor="middle" font-size="10" fill="#8f8f8f">${esc(p.label.length > room ? p.label.slice(0, room - 1) + '…' : p.label)}</text>`;
      return `<g><title>${esc((p.label ? p.label + ': ' : '') + p.text)}</title>
        <rect x="${(cx - bw / 2).toFixed(1)}" y="${top.toFixed(1)}" width="${bw.toFixed(1)}" height="${h.toFixed(1)}" fill="${p.v < 0 ? '#ff4b5c' : '#45a873'}"/>${labels}</g>`;
    }).join('');
    return `<div class="wb-chart-h"><b>${esc(wbLabel(row) || 'Row ' + (k + 1))}</b><span class="muted">${esc(sheet.name)} · click another row to chart it</span></div>
      <svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Chart of ${esc(wbLabel(row))}">
        <line x1="${L}" x2="${W - R}" y1="${y(0).toFixed(1)}" y2="${y(0).toFixed(1)}" stroke="#474747"/>${bars}</svg>`;
  }
  function wbView(i) {
    const w = i.workbook, sheet = w.sheets[state.wb.sheet] || w.sheets[0];
    const cols = Math.max(...sheet.rows.map((r) => r.length));
    const tabs = w.sheets.map((s, n) => `<button class="tab${s === sheet ? ' on' : ''}" data-wbsheet="${n}">${esc(s.name)}</button>`).join('');
    const body = sheet.rows.map((row, k) => (row.every((c) => c === 0)
      ? `<tr class="gap"><td colspan="${cols}"></td></tr>` : wbRowHtml(sheet, row, k))).join('');
    const ext = w.file ? w.file.slice(w.file.lastIndexOf('.')) : '';
    return `<div class="model-bar"><b>${i.ticker} model</b><span class="muted">${esc(w.name)} · saved ${fmtDate(w.updated)}</span><span class="sp"></span>
        ${w.file ? `<a class="btn" href="${esc(w.file)}" download="${esc(w.name)}">Download ${esc(ext)}</a>` : ''}</div>
      <div id="wbChart">${state.wb.row != null ? wbChart(sheet, state.wb.row)
        : sheet.rows.some((r, n) => wbChartable(r) && n !== sheet.hdr) ? '<p class="note">Click a row of numbers to chart it.</p>' : ''}</div>
      <div class="tabs wb-tabs">${tabs}</div>
      <div class="wb-scroll"><table class="grid wb"><tbody>${body}</tbody></table></div>
      ${sheet.truncated ? '<p class="note">This sheet is very large; only its first rows and columns are shown. The download has all of it.</p>' : ''}`;
  }
  // Clicks on a sheet tab or a row of the Excel model.
  function wbClick(d) {
    const i = ideaBy[state.wb.ticker];
    if (!i || !i.workbook) return;
    if (d.wbsheet != null) {
      state.wb.sheet = +d.wbsheet;
      state.wb.row = wbDefaultRow(i.workbook.sheets[state.wb.sheet]);
      $('#wbView').innerHTML = wbView(i);
    } else {
      state.wb.row = +d.wbrow;
      $('#wbChart').innerHTML = wbChart(i.workbook.sheets[state.wb.sheet], state.wb.row);
      document.querySelectorAll('#wbView tr.pick').forEach((tr) => tr.classList.toggle('sel', +tr.dataset.wbrow === state.wb.row));
    }
  }

  // ---------- reader for one update ----------
  const fontButtons = () => `<span class="fs">${[0, 1, 2].map((v) => `<button data-act="font" data-v="${v}" class="${state.font === v ? 'on' : ''}" style="font-size:${10 + v * 2}px" title="Text size">A</button>`).join('')}</span>`;
  function viewPost(p) {
    return `<article class="reader f${state.font}">
      <div class="reader-bar"><button data-act="back">« Back</button><button data-act="copy">Copy link</button><span class="sp"></span>${fontButtons()}<button data-act="print">Export PDF</button></div>
      <div class="meta"><span class="type t-${p.type}">${p.type}</span> · ${fmtDate(p.date).toUpperCase()} ${p.time}</div>
      <h1>${esc(p.headline)}</h1>
      ${p.tickers.length ? `<div class="matched">Matched: ${p.tickers.map((t) => `<button class="tkchip" data-go="#/s/${t}">${t}</button>`).join('')}</div>` : ''}
      ${p.change ? `<div class="change">Tracker change · ${esc(p.change.field)}: <s>${esc(p.change.from)}</s> → <b>${esc(p.change.to)}</b></div>` : ''}
      ${p.body ? p.body.map((x) => `<p>${esc(x)}</p>`).join('') : '<p class="muted">Headline-only update.</p>'}
      ${p.link ? `<p><button class="btn" data-go="${p.link.go}">${esc(p.link.label)} ›</button></p>` : ''}
    </article>`;
  }

  // ---------- profile pages ----------
  const cmdChip = (code) => `<button class="chip" data-cmd="${code}">${code}</button>`;
  // Each row is [when, what, detail, bullet points]; the last two are optional.
  const rowsTable = (rows) => `<table class="rows">${rows.map(([w, t, d, pts]) => `<tr><td>${esc(w)}</td><td><b>${esc(t)}</b>${d ? `<div class="d">${esc(d)}</div>` : ''}${pts && pts.length ? `<ul class="pts">${pts.map((p) => `<li>${esc(p)}</li>`).join('')}</ul>` : ''}</td></tr>`).join('')}</table>`;

  function trackRecord() {
    const calls = D.ideas.filter((i) => calc(i).ret != null);
    const rets = calls.map((i) => calc(i).ret);
    const avg = rets.reduce((a, b) => a + b, 0) / rets.length;
    const wins = rets.filter((r) => r > 0).length;
    const cp = cpTally(D.ideas.flatMap((i) => i.checkpoints || []));
    return `<div class="statrow">
      <div class="stat"><b>${D.ideas.filter((i) => i.status === 'active').length}</b><span>active calls</span></div>
      <div class="stat"><b>${D.ideas.filter((i) => i.status === 'closed').length}</b><span>closed calls</span></div>
      <div class="stat"><b>${wins}/${calls.length}</b><span>in the money</span></div>
      <div class="stat"><b class="${cls(avg)}">${pct(avg)}</b><span>average return</span></div>
      ${cp.settled ? `<div class="stat"><b>${cp.good}/${cp.settled}</b><span>checkpoints went my way</span></div>` : ''}
    </div>`;
  }
  function viewAbout() {
    const o = D.owner;
    return `<div class="page">
      <div class="hero"><div class="avatar">${esc(o.initials)}</div><div><h1>${esc(o.name)}</h1><div class="muted">${esc(o.role)}</div></div></div>
      <p class="lead">${esc(o.tagline)}</p>
      <div class="callout">New here? Click any row in the quote monitor to open the thesis behind it.<br>
        Or run a command: ${cmdChip('QM')} ${cmdChip('N')}${anyCheckpoints ? ' ' + cmdChip('CP') : ''} ${cmdChip('CONTACT')} ${cmdChip('HELP')}</div>
      <h2>Track record</h2>${trackRecord()}
      <p class="note" style="margin-top:8px">Calculated from the monitor, closed calls included.${D.hasSamples ? ' Includes sample data.' : ''}</p>
      ${o.education.length ? `<h2>Education</h2>${rowsTable(o.education)}` : ''}
      ${o.experience.length ? `<h2>Experience</h2>${rowsTable(o.experience)}` : ''}
      ${o.skills.length ? `<h2>Skills</h2><p>${o.skills.map((s) => `<span class="tkchip">${esc(s)}</span>`).join(' ')}</p>` : ''}
    </div>`;
  }

  function viewContact() {
    const o = D.owner;
    const a = (href, text) => `<a href="${esc(href)}"${/^https?:/.test(href) ? ' target="_blank" rel="noopener"' : ''}>${esc(text)}</a>`;
    const web = (addr) => (/^https?:\/\//i.test(addr) ? addr : 'https://' + addr);
    const rows = [];
    if (o.email) rows.push(['Email', a('mailto:' + o.email, o.email)]);
    if (o.phone) rows.push(['Phone', a('tel:' + o.phone.replace(/[^\d+]/g, ''), o.phone)]);
    o.links.forEach(([label, addr]) => rows.push([label, a(web(addr), addr)]));
    return `<div class="page"><p class="lead">Open to roles in equity research and investing.</p>
      <table class="rows">${rows.map(([k, v]) => `<tr><td>${esc(k)}</td><td>${v}</td></tr>`).join('')}</table></div>`;
  }
  function viewHelp() {
    const g = GLOBAL.map((c) => `<tr><td class="l">${cmdChip(c.code)}</td><td class="l">${esc(c.desc)}</td></tr>`).join('');
    const t = TICKER.map((c) => `<tr><td class="l"><span class="chip tk">HLXB</span> <span class="chip">${c.code}</span></td><td class="l">${esc(c.desc)}</td></tr>`).join('');
    return `<div class="page"><h1>Help</h1>
      <p class="lead">Everything here works by clicking. The command bar is the faster way round.</p>
      <h2>Commands</h2><table class="grid" style="white-space:normal"><tbody>${g}</tbody></table>
      <h2>Ticker commands</h2><p class="muted">Type a ticker, then a command. For example <b>HLXB MODEL</b>.</p>
      <table class="grid" style="white-space:normal"><tbody>${t}</tbody></table>
      <h2>Keys</h2>${kv([['` or /', 'Focus the command bar'], ['↑ ↓', 'Move through results'], ['Enter', 'Run the highlighted result'], ['Esc', 'Close the results'], ['Backspace on empty bar', 'Clear the loaded ticker']])}
      <h2>Windows</h2>${kv([['Drag a title bar', 'Move the window'], ['Drag an edge or corner', 'Resize it; windows sharing that edge follow'], ['Alt + drag an edge', 'Resize one window on its own'], ['Double-click a title bar, or ⤢', 'Maximise or restore'], ['×', 'Close; QM, N, CONTACT or any page command reopens it'], ['Reset layout', 'Back to the default arrangement']])}
      <h2>About the data</h2><p class="muted">Prices are entered by hand and dated, not live. Views are personal and not investment advice.</p></div>`;
  }

  // ---------- panels, screens, routing ----------
  // Each window's position and size is stored as fractions of the workspace, so
  // a layout survives a browser resize. Screens (bottom tabs) are presets.
  const wsEl = $('#workspace');
  const PANELS = ['p-qm', 'p-wire', 'p-main', 'p-contact'];
  const GAP = 3, MIN_W = 260, MIN_H = 150, SNAP = 8, STORE = 'pt-layout-v1';
  const narrow = () => matchMedia('(max-width: 720px)').matches;
  const rect = (x, y, w, h, open = true) => ({ x, y, w, h, open });
  let zTop = 10;

  function preset(name) {
    const main = rect(0.52, 0, 0.48, 1);
    // Contact starts closed, except on the About screen. The CONTACT command opens
    // it in the lower left, beside the main window rather than over it.
    const contact = rect(0.17, 0.6, 0.35, 0.4, false);
    if (name === 'Tracker') return { 'p-qm': rect(0, 0, 0.52, 1), 'p-wire': rect(0, 0.5, 0.52, 0.5, false), 'p-main': main, 'p-contact': contact };
    if (name === 'Wire') return { 'p-qm': rect(0, 0, 0.52, 0.5, false), 'p-wire': rect(0, 0, 0.52, 1), 'p-main': main, 'p-contact': contact };
    if (name === 'About') return { 'p-qm': rect(0, 0, 0.52, 0.5, false), 'p-wire': rect(0, 0.5, 0.52, 0.5, false), 'p-main': rect(0, 0, 0.66, 1), 'p-contact': rect(0.66, 0, 0.34, 1) };
    return wsEl.clientWidth >= 1100
      ? { 'p-qm': rect(0, 0, 0.52, 0.5), 'p-wire': rect(0, 0.5, 0.52, 0.5), 'p-main': main, 'p-contact': contact }
      : { 'p-qm': rect(0, 0, 1, 0.4), 'p-wire': rect(0, 0.4, 0.45, 0.6), 'p-main': rect(0.45, 0.4, 0.55, 0.6), 'p-contact': rect(0, 0.6, 0.45, 0.4, false) };
  }
  function applyLayout() {
    const W = wsEl.clientWidth, H = wsEl.clientHeight;
    let any = false;
    PANELS.forEach((id) => {
      const p = $('#' + id), g = state.layout[id];
      const r = g.max ? { x: 0, y: 0, w: 1, h: 1 } : g;
      p.hidden = !g.open;
      any = any || g.open;
      p.style.left = r.x * W + GAP + 'px';
      p.style.top = r.y * H + GAP + 'px';
      p.style.width = Math.max(0, r.w * W - 2 * GAP) + 'px';
      p.style.height = Math.max(0, r.h * H - 2 * GAP) + 'px';
    });
    $('#wsEmpty').hidden = any;
  }
  function saveLayout() {
    try { localStorage.setItem(STORE, JSON.stringify({ ws: state.ws, layout: state.layout })); } catch (_) { /* storage unavailable */ }
  }
  function renderWsTabs() {
    $('#wsTabs').innerHTML = ['Main', 'Tracker', 'Wire', 'About'].map((n) => `<button class="${n === state.ws ? 'on' : ''}" data-ws="${n}">${n}</button>`).join('');
  }
  function setWs(name) {
    state.ws = name;
    state.layout = preset(name);
    applyLayout();
    renderWsTabs();
    saveLayout();
  }
  // The layout no longer matches a preset once the visitor changes it by hand.
  function layoutChanged() {
    state.ws = null;
    applyLayout();
    renderWsTabs();
    saveLayout();
  }
  function initLayout() {
    let saved = null;
    try { saved = JSON.parse(localStorage.getItem(STORE)); } catch (_) { /* ignore */ }
    // A layout saved before the Contact window existed keeps its other windows as they were.
    if (saved && saved.layout && !saved.layout['p-contact']) saved.layout['p-contact'] = preset('Main')['p-contact'];
    const ok = saved && saved.layout && PANELS.every((id) => {
      const g = saved.layout[id];
      return g && ['x', 'y', 'w', 'h'].every((k) => typeof g[k] === 'number' && g[k] >= 0 && g[k] <= 1);
    });
    if (!ok) return setWs('Main');
    state.ws = saved.ws || null;
    state.layout = saved.layout;
    applyLayout();
    renderWsTabs();
  }
  function raise(p) {
    p.style.zIndex = ++zTop;
    document.querySelectorAll('.panel').forEach((x) => x.classList.toggle('focus', x === p));
  }
  function focusPanel(id) {
    const p = $('#' + id);
    if (!state.layout[id].open) { state.layout[id].open = true; layoutChanged(); }
    raise(p);
    if (booted && narrow()) p.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  function closePanel(id) { state.layout[id].open = false; state.layout[id].max = false; layoutChanged(); }
  function toggleMax(id) { state.layout[id].max = !state.layout[id].max; raise($('#' + id)); layoutChanged(); }

  // Drag a header to move a window, or an edge or corner to resize it. The moving
  // edges snap to the workspace edges and to the edges of the other open windows.
  function startDrag(e, p, mode) {
    if (e.button !== 0 || narrow()) return;
    const g = state.layout[p.id], W = wsEl.clientWidth, H = wsEl.clientHeight;
    if (g.max) Object.assign(g, { x: 0, y: 0, w: 1, h: 1, max: false });
    const s = { x: g.x * W, y: g.y * H, w: g.w * W, h: g.h * H }, px0 = e.clientX, py0 = e.clientY;
    const xs = [0, W], ys = [0, H];
    PANELS.forEach((id) => {
      const q = state.layout[id];
      if (id !== p.id && q.open && !q.max) { xs.push(q.x * W, (q.x + q.w) * W); ys.push(q.y * H, (q.y + q.h) * H); }
    });
    const snap = (v, lines) => { for (const l of lines) if (Math.abs(v - l) <= SNAP) return l; return v; };
    const clamp = (v, lo, hi) => Math.min(Math.max(v, lo), Math.max(lo, hi));

    // A resized edge acts as a divider: every window edge that is flush with it,
    // on either side, moves with it, so no gap or overlap opens up between
    // neighbours. Hold Alt to resize this window on its own.
    const others = PANELS.filter((id) => id !== p.id && state.layout[id].open && !state.layout[id].max)
      .map((id) => { const q = state.layout[id]; return { g: q, x: q.x * W, y: q.y * H, w: q.w * W, h: q.h * H }; });
    const self = { g, ...s };
    const divider = (pos, len, cross, crossLen, end, minLen, total, lines) => {
      const line0 = end ? self[pos] + self[len] : self[pos];
      const group = [{ r: self, end }];
      if (!e.altKey) {
        const flush = others.map((r) => (Math.abs(r[pos] + r[len] - line0) <= 2 ? { r, end: true }
          : Math.abs(r[pos] - line0) <= 2 ? { r, end: false } : null)).filter(Boolean);
        // Only windows that touch the group along the divider join it.
        const touches = (a, b) => a[cross] <= b[cross] + b[crossLen] + 2 && b[cross] <= a[cross] + a[crossLen] + 2;
        for (let grew = true; grew;) {
          grew = false;
          flush.forEach((c) => {
            if (!group.includes(c) && group.some((m) => touches(m.r, c.r))) { group.push(c); grew = true; }
          });
        }
        // With no window on the far side this is not a divider, so nothing else follows.
        if (!group.some((m) => m.end !== end)) group.length = 1;
      }
      let lo = 0, hi = total;
      group.forEach(({ r, end: isEnd }) => {
        if (isEnd) lo = Math.max(lo, r[pos] + minLen); else hi = Math.min(hi, r[pos] + r[len] - minLen);
      });
      return (delta, next) => {
        const line = clamp(snap(line0 + delta, lines), lo, hi);
        group.forEach(({ r, end: isEnd }) => {
          const n = next.get(r);
          if (isEnd) n[len] = line - r[pos]; else { n[pos] = line; n[len] = r[pos] + r[len] - line; }
        });
      };
    };
    const moveX = mode === 'move' ? null
      : mode.includes('e') ? divider('x', 'w', 'y', 'h', true, MIN_W, W, xs)
      : mode.includes('w') ? divider('x', 'w', 'y', 'h', false, MIN_W, W, xs) : null;
    const moveY = mode === 'move' ? null
      : mode.includes('s') ? divider('y', 'h', 'x', 'w', true, MIN_H, H, ys)
      : mode.includes('n') ? divider('y', 'h', 'x', 'w', false, MIN_H, H, ys) : null;

    let moved = false;
    const move = (ev) => {
      const dx = ev.clientX - px0, dy = ev.clientY - py0;
      if (!moved && Math.abs(dx) + Math.abs(dy) < 3) return;
      moved = true;
      if (mode === 'move') {
        let { x, y } = s;
        const { w, h } = s;
        x = s.x + dx; y = s.y + dy;
        const l = snap(x, xs), r = snap(x + w, xs), t = snap(y, ys), b = snap(y + h, ys);
        x = clamp(l !== x ? l : r !== x + w ? r - w : x, 0, W - w);
        y = clamp(t !== y ? t : b !== y + h ? b - h : y, 0, H - h);
        Object.assign(g, { x: x / W, y: y / H });
      } else {
        const next = new Map([self, ...others].map((r) => [r, { x: r.x, y: r.y, w: r.w, h: r.h }]));
        if (moveX) moveX(dx, next);
        if (moveY) moveY(dy, next);
        next.forEach((n, r) => Object.assign(r.g, { x: n.x / W, y: n.y / H, w: n.w / W, h: n.h / H }));
      }
      applyLayout();
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      window.removeEventListener('pointercancel', up);
      document.body.classList.remove('dragging');
      if (moved) layoutChanged();
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
    window.addEventListener('pointercancel', up);
    document.body.classList.add('dragging');
    raise(p);
  }
  PANELS.forEach((id) => {
    const p = $('#' + id);
    ['n', 's', 'e', 'w', 'ne', 'nw', 'se', 'sw'].forEach((d) => {
      const h = document.createElement('div');
      h.className = 'rz rz-' + d;
      h.addEventListener('pointerdown', (e) => { e.preventDefault(); startDrag(e, p, d); });
      p.appendChild(h);
    });
    const head = $('.panel-head', p);
    head.addEventListener('pointerdown', (e) => { if (!e.target.closest('button')) startDrag(e, p, 'move'); });
    head.addEventListener('dblclick', (e) => { if (!e.target.closest('button') && !narrow()) toggleMax(id); });
  });
  new ResizeObserver(applyLayout).observe(wsEl);
  function setMain(title, sub, html) {
    $('#mainTitle').textContent = title;
    $('#mainSub').textContent = sub;
    const b = $('#mainBody');
    b.innerHTML = html;
    b.scrollTop = 0;
    drawCharts();
  }
  // A chart is drawn to its pixel width, so it is redrawn when the window's width changes.
  let chartW = 0, chartRaf = 0;
  new ResizeObserver(() => {
    const w = $('#mainBody').clientWidth;
    if (w === chartW) return;
    chartW = w;
    cancelAnimationFrame(chartRaf);
    chartRaf = requestAnimationFrame(drawCharts);
  }).observe($('#mainBody'));
  function setCtx(t) {
    state.ctx = t;
    const c = $('#ctx');
    c.hidden = !t;
    c.innerHTML = t ? `${t}<span class="x">×</span>` : '';
    renderQM();
  }
  function go(hash) {
    if (location.hash === hash) route(); else location.hash = hash;
  }
  function route() {
    const [a, b, c] = location.hash.replace(/^#\/?/, '').split('/').filter(Boolean).map(decodeURIComponent);
    const i = a === 's' ? ideaBy[(b || '').toUpperCase()] : null;
    if (i) {
      setCtx(i.ticker);
      if (c === 'thesis') setMain('THESIS', i.ticker, secNav(i, 'thesis') + viewThesis(i));
      else if (c === 'model') setMain('MODEL', i.ticker, secNav(i, 'model') + viewModel(i));
      else if (c === 'cp') setMain('CHECKPOINTS', i.ticker, secNav(i, 'cp') + viewCheckpoints(i.ticker));
      else {
        setMain('DESCRIPTION', i.ticker, secNav(i, 'des') + viewDes(i));
        if (c === 'n') { state.wire.ticker = i.ticker; renderWire(); return focusPanel('p-wire'); }
      }
    } else if (a === 'n' && postBy[b]) setMain('UPDATE', postBy[b].tickers.join(' '), viewPost(postBy[b]));
    else if (a === 'help') setMain('HELP', '', viewHelp());
    else if (a === 'cp') setMain('CHECKPOINTS', '', viewCheckpoints());
    else setMain('ABOUT', '', viewAbout());
    focusPanel('p-main');
    if (a === 'contact') focusPanel('p-contact');  // an old link to the contact page
  }

  // ---------- commands and palette ----------
  const GLOBAL = [
    { code: 'ABOUT', desc: 'Who I am and what this terminal is', run: () => go('#/about') },
    { code: 'QM', desc: 'Quote monitor: every idea I track, with my call and the return since', run: () => focusPanel('p-qm') },
    { code: 'N', desc: 'Updates wire: thesis changes, model revisions and notes', run: () => { clearWire(); focusPanel('p-wire'); } },
    { code: 'CONTACT', desc: 'Email, phone and profiles, in a window of their own', run: () => focusPanel('p-contact') },
    { code: 'HELP', desc: 'Commands and keyboard shortcuts', run: () => go('#/help') },
  ];
  // Offered once a ticker note has a Checkpoints section.
  if (anyCheckpoints) GLOBAL.splice(3, 0, { code: 'CP', desc: 'Checkpoints: what each call needs next, soonest first', run: () => go('#/cp') });
  const TICKER = [
    { code: 'DES', desc: 'Overview: my call, thesis in brief, checkpoints, chart and catalysts', sub: '' },
    { code: 'THESIS', desc: 'The full written thesis', sub: '/thesis' },
    { code: 'MODEL', desc: 'Financial model and scenario valuation', sub: '/model' },
    { code: 'CP', desc: 'Checkpoints for this name: what the call needs, and what would close it', sub: '/cp' },
    { code: 'N', desc: 'Updates on this name', sub: '/n' },
  ];
  const input = $('#cmdInput'), pal = $('#palette');

  function buildResults(raw) {
    const q = raw.trim().toUpperCase(), tokens = q.split(/\s+/).filter(Boolean);
    const typedTicker = tokens.length && ideaBy[tokens[0]] ? tokens[0] : null;
    const tk = typedTicker || state.ctx;
    const cq = (typedTicker ? tokens.slice(1) : tokens).join(' ');
    const match = (c) => !cq || c.code.startsWith(cq) || c.desc.toUpperCase().includes(cq);
    const rank = (a, b) => (b.code.startsWith(cq) ? 1 : 0) - (a.code.startsWith(cq) ? 1 : 0);
    const items = [];
    if (tk) {
      // CP is offered only for a ticker that has checkpoints.
      TICKER.filter((c) => c.code !== 'CP' || ideaBy[tk].checkpoints).filter(match).sort(rank).forEach((c) => items.push({
        group: 'COMMANDS',
        html: `<span class="chip tk">${tk}</span><span class="chip">${c.code}</span><span class="desc">${esc(c.desc)}</span>`,
        run: () => go(`#/s/${tk}${c.sub}`),
      }));
    }
    if (!typedTicker) {
      GLOBAL.filter(match).sort(rank).forEach((c) => items.push({
        group: 'COMMANDS', html: `<span class="chip">${c.code}</span><span class="desc">${esc(c.desc)}</span>`, run: c.run,
      }));
      D.ideas.filter((i) => !q || i.ticker.startsWith(q) || i.name.toUpperCase().includes(q)).forEach((i) => {
        const c = calc(i);
        items.push({
          group: 'IDEAS',
          html: `<span class="chip tk">${i.ticker}</span><span class="badge ${i.stance.toLowerCase()}">${i.stance}</span><span class="desc">${esc(i.name)} <span class="muted">· ${esc(i.sector)}</span></span><span class="${cls(c.ret)}">${c.ret == null ? '' : pct(c.ret)}</span>`,
          run: () => go(`#/s/${i.ticker}`),
        });
      });
      if (q.length >= 2) {
        posts.filter((p) => p.headline.toUpperCase().includes(q)).slice(0, 5).forEach((p) => items.push({
          group: 'UPDATES',
          html: `<span class="type t-${p.type}">${p.type}</span><span class="desc">${esc(p.headline)}</span><span class="muted">${fmtDate(p.date)}</span>`,
          run: () => go(`#/n/${p.id}`),
        }));
      }
    }
    // An exact command or ticker beats everything else when Enter is pressed.
    items.sort((a, b) => (a.group === b.group ? 0 : ['COMMANDS', 'IDEAS', 'UPDATES'].indexOf(a.group) - ['COMMANDS', 'IDEAS', 'UPDATES'].indexOf(b.group)));
    return items;
  }
  function renderPal() {
    const items = state.pal.items;
    let html = '', g = null;
    items.forEach((it, k) => {
      if (it.group !== g) { g = it.group; html += `<div class="pal-g">${g}</div>`; }
      const on = k === state.pal.idx;
      html += `<button class="pal-item${on ? ' on' : ''}" data-pal="${k}" tabindex="-1">${it.html}${on ? '<span class="hint">Enter ↵</span>' : ''}</button>`;
    });
    pal.innerHTML = html || '<div class="pal-empty">No matches. Try HELP.</div>';
    pal.hidden = false;
    const on = $('.pal-item.on', pal);
    if (on) on.scrollIntoView({ block: 'nearest' });
  }
  function openPal() {
    state.pal.items = buildResults(input.value);
    // Prefer an idea whose ticker was typed exactly over a command that merely matches.
    const q = input.value.trim().toUpperCase();
    const exact = q ? state.pal.items.findIndex((it) => it.group === 'IDEAS' && it.html.includes(`>${q}<`)) : -1;
    const cmdExact = GLOBAL.some((c) => c.code === q) || TICKER.some((c) => c.code === q);
    state.pal.idx = exact >= 0 && !cmdExact ? exact : 0;
    renderPal();
  }
  function closePal() { pal.hidden = true; }
  function runPal(k) {
    const it = state.pal.items[k];
    if (!it) return;
    input.value = '';
    closePal();
    input.blur();
    it.run();
  }

  input.addEventListener('focus', openPal);
  input.addEventListener('click', openPal);
  input.addEventListener('input', openPal);
  input.addEventListener('keydown', (e) => {
    const n = state.pal.items.length;
    if (e.key === '`') e.preventDefault();
    else if (e.key === 'ArrowDown' && n) { e.preventDefault(); state.pal.idx = (state.pal.idx + 1) % n; renderPal(); }
    else if (e.key === 'ArrowUp' && n) { e.preventDefault(); state.pal.idx = (state.pal.idx - 1 + n) % n; renderPal(); }
    else if (e.key === 'Enter') { e.preventDefault(); runPal(state.pal.idx); }
    else if (e.key === 'Backspace' && !input.value && state.ctx) { setCtx(null); openPal(); }
  });
  // ---------- picture zoom ----------
  // A picture in a thesis opens enlarged over the page. A click anywhere, or Esc, closes it.
  const closeZoom = () => { const z = $('#zoom'); if (z) z.remove(); };
  const openZoom = (pic) => {
    closeZoom();
    const small = pic.querySelector('img');
    const figure = pic.closest('figure');
    const cap = figure && figure.querySelector('figcaption');
    const z = document.createElement('div');
    z.id = 'zoom';
    z.setAttribute('role', 'dialog');
    z.setAttribute('aria-label', 'Enlarged picture');
    z.innerHTML = `<button class="zoom-x" aria-label="Close">×</button>
      <img src="${esc(pic.getAttribute('href'))}" alt="${esc(small ? small.alt : '')}">
      ${cap ? `<div class="zoom-cap">${esc(cap.textContent)}</div>` : ''}`;
    const big = z.querySelector('img');
    // The picture grows to fill the window, up to twice its own size.
    const size = () => {
      z.style.setProperty('--nw', big.naturalWidth);
      z.style.setProperty('--ar', big.naturalWidth / big.naturalHeight || 1);
    };
    big.addEventListener('load', size);
    if (big.complete && big.naturalWidth) size();
    z.addEventListener('click', closeZoom);
    document.body.appendChild(z);
    z.querySelector('.zoom-x').focus();
  };
  window.addEventListener('hashchange', closeZoom);

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && $('#zoom')) { closeZoom(); return; }
    const inField = /INPUT|SELECT|TEXTAREA/.test(document.activeElement.tagName);
    if ((e.key === '`' || e.key === '/') && !inField) { e.preventDefault(); input.focus(); openPal(); }
    else if (e.key === 'Escape') { closePal(); input.blur(); }
    else if (e.key === 'Enter' && document.activeElement.matches('tr[data-go]')) go(document.activeElement.dataset.go);
    else if (e.key === 'Enter' && document.activeElement.matches('tr[data-wbrow]')) document.activeElement.click();
  });
  document.addEventListener('pointerdown', (e) => {
    if (!e.target.closest('.topbar')) closePal();
    const p = e.target.closest('.panel');
    if (p) raise(p);
  });

  // ---------- clicks ----------
  document.addEventListener('click', (e) => {
    // A thesis picture zooms in place. Ctrl, Cmd or Shift with the click still opens it in a new tab.
    const pic = e.target.closest('.md a.pic');
    if (pic && !(e.ctrlKey || e.metaKey || e.shiftKey)) { e.preventDefault(); return openZoom(pic); }
    const t = e.target.closest('[data-pal],[data-go],[data-cmd],[data-act],[data-qmtab],[data-sort],[data-ws],[data-wbsheet],[data-wbrow]');
    if (!t) return;
    const d = t.dataset;
    if (d.pal != null) return runPal(+d.pal);
    if (d.wbsheet != null || d.wbrow != null) return wbClick(d);
    if (d.go) return go(d.go);
    if (d.cmd) return GLOBAL.find((c) => c.code === d.cmd).run();
    if (d.qmtab) { state.tab = d.qmtab; return renderQM(); }
    if (d.sort) {
      state.sort = { k: d.sort, dir: state.sort.k === d.sort ? -state.sort.dir : 1 };
      return renderQM();
    }
    if (d.ws) { setWs(d.ws); if (d.ws === 'About') go('#/about'); return; }
    if (d.act === 'max') return toggleMax(t.closest('.panel').id);
    if (d.act === 'close') return closePanel(t.closest('.panel').id);
    if (d.act === 'reset') return setWs('Main');
    if (d.act === 'back') return history.length > 1 ? history.back() : go('#/about');
    if (d.act === 'print') return window.print();
    if (d.act === 'range') { state.range = d.v; return drawCharts(); }
    if (d.act === 'banner-x') { $('#banner').hidden = true; return; }
    if (d.act === 'uifont') return setUiFont(+d.v, true);
    if (d.act === 'font') {
      state.font = +d.v;
      const r = $('.reader');
      r.className = 'reader f' + state.font;
      r.querySelectorAll('.fs button').forEach((b) => b.classList.toggle('on', +b.dataset.v === state.font));
      return;
    }
    if (d.act === 'copy') {
      const done = () => { t.textContent = 'Link copied'; setTimeout(() => (t.textContent = 'Copy link'), 1500); };
      try { navigator.clipboard.writeText(location.href).then(done, done); } catch (_) { done(); }
    }
  });
  $('#ctx').addEventListener('click', () => { setCtx(null); input.focus(); });
  $('#qmQ').addEventListener('input', (e) => { state.qmQ = e.target.value; renderQM(); });
  $('#qmQ').addEventListener('keydown', (e) => {
    const first = $('#qmTable tbody tr[data-go]');
    if (e.key === 'Enter' && first) go(first.dataset.go);
    else if (e.key === 'Escape' && state.qmQ) { e.stopPropagation(); e.target.value = state.qmQ = ''; renderQM(); }
  });
  $('#wireQ').addEventListener('input', (e) => { state.wire.q = e.target.value; renderWire(); });
  $('#wireType').addEventListener('change', (e) => { state.wire.type = e.target.value; renderWire(); });
  $('#wireTk').addEventListener('click', () => { state.wire.ticker = null; renderWire(); });
  $('#bannerText').addEventListener('click', () => go(`#/n/${posts[0].id}`));

  // ---------- text size (the AAA at the right end of the status bar) ----------
  // Sets --ui-scale, which style.css applies to the top bar and to the contents
  // of every window. The choice is remembered in this browser.
  // 1 is the size the stylesheet was written at (13px body text); it is the largest step.
  const UI_SCALES = [0.8, 0.9, 1], UI_DEFAULT = 1, UI_STORE = 'pt-text-size-v2';
  function setUiFont(v, save) {
    if (!(v in UI_SCALES)) v = UI_DEFAULT;
    document.documentElement.style.setProperty('--ui-scale', UI_SCALES[v]);
    document.querySelectorAll('.ui-fs button').forEach((b) => {
      b.classList.toggle('on', +b.dataset.v === v);
      b.setAttribute('aria-pressed', +b.dataset.v === v);
    });
    if (save) try { localStorage.setItem(UI_STORE, String(v)); } catch (_) { /* storage unavailable */ }
  }
  let savedSize = UI_DEFAULT;
  try { const raw = localStorage.getItem(UI_STORE); if (raw !== null) savedSize = +raw; } catch (_) { /* ignore */ }
  setUiFont(savedSize, false);

  // ---------- boot ----------
  $('#pins').innerHTML = ['QM', 'N', 'ABOUT', 'HELP'].map(cmdChip).join('');
  $('#contactBody').innerHTML = viewContact();
  $('#statusNote').textContent = `${D.hasSamples ? 'INCLUDES SAMPLE DATA · ' : ''}PRICES AS OF ${fmtDate(D.asOf).toUpperCase()} · PERSONAL VIEWS, NOT INVESTMENT ADVICE · v0.1`;
  $('#bannerText').textContent = `LATEST · ${fmtDate(posts[0].date).toUpperCase()} | ${posts[0].headline}`;
  $('#banner').hidden = false;
  const tick = () => { $('#clock').textContent = new Date().toLocaleTimeString('en-GB'); };
  tick();
  setInterval(tick, 1000);

  initLayout();
  renderQM();
  renderWire();
  window.addEventListener('hashchange', route);
  route();
  booted = true;
})();
