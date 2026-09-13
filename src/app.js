(function () {
      var LIBS = [
        {
          name: 'XLSX', urls: [
            'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
            'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js',
            'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
          ]
        },
        {
          name: 'echarts', urls: [
            'https://cdn.jsdelivr.net/npm/echarts@5.5.1/dist/echarts.min.js',
            'https://unpkg.com/echarts@5.5.1/dist/echarts.min.js',
            'https://cdnjs.cloudflare.com/ajax/libs/echarts/5.5.1/echarts.min.js'
          ]
        }
      ];
      function loadOne(lib, i) {
        return new Promise(function (resolve) {
          if (window[lib.name]) return resolve(true);
          if (i >= lib.urls.length) return resolve(false);
          var s = document.createElement('script');
          s.src = lib.urls[i]; s.async = true;
          s.onload = function () { resolve(!!window[lib.name]); };
          s.onerror = function () { s.remove(); loadOne(lib, i + 1).then(resolve); };
          document.head.appendChild(s);
        });
      }
      window.__libsReady = Promise.all(LIBS.map(function (l) { return loadOne(l, 0); }))
        .then(function (r) { window.__libStatus = { XLSX: r[0], echarts: r[1] }; return window.__libStatus; });
    })();

/* ==UTILS-START== */
    'use strict';
    const $ = (s, r) => (r || document).querySelector(s);
    const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
    const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
    const uid = () => Math.random().toString(36).slice(2, 10);

    function esc(v) {
      if (v === null || v === undefined) return '';
      return String(v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    }

    function isBlank(v) {
      if (v === null || v === undefined) return true;
      if (typeof v === 'number') return !isFinite(v);
      if (v instanceof Date) return isNaN(v.getTime());
      return String(v).trim() === '' || String(v).toLowerCase() === 'nan' || String(v).toLowerCase() === 'null' || String(v).toLowerCase() === 'undefined';
    }

    const UNIT_MAP = { '万': 1e4, '亿': 1e8, '千': 1e3, 'w': 1e3, 'W': 1e3, 'k': 1e3, 'K': 1e3, 'm': 1e6, 'M': 1e6, 'b': 1e9, 'B': 1e9 };

    /** 宽容的数字解析：支持千分位、货币符号、百分号、括号负数、中文单位 */
    function toNum(v) {
      if (v === null || v === undefined) return NaN;
      if (typeof v === 'number') return isFinite(v) ? v : NaN;
      if (typeof v === 'boolean') return v ? 1 : 0;
      if (v instanceof Date) { const t = v.getTime(); return isNaN(t) ? NaN : t; }
      let s = String(v).trim();
      if (!s) return NaN;
      let mult = 1;
      if (/%$/.test(s)) { mult = 0.01; s = s.slice(0, -1).trim(); }
      s = s.replace(/[¥￥$€£,\s ]/g, '');
      if (/^\(.*\)$/.test(s)) s = '-' + s.slice(1, -1);
      const um = s.match(/^(-?[\d.]+)([万亿千wWkKmMbB])$/);
      if (um && UNIT_MAP[um[2]]) { s = um[1]; mult *= UNIT_MAP[um[2]]; }
      if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) {
        const lead = s.match(/^[+-]?(\d+\.?\d*|\.\d+)/);
        if (!lead) return NaN;
        s = lead[0];
      }
      const n = parseFloat(s) * mult;
      return isFinite(n) ? n : NaN;
    }

    function looksLikeDate(s) {
      return /^\d{4}[-/.年]\d{1,2}([-/.月]\d{1,2}日?)?/.test(s)
        || /^\d{1,2}[-/.]\d{1,2}[-/.]\d{4}/.test(s)
        || /^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}/.test(s)
        || /^\d{8}$/.test(s) && false;
    }

    /** 宽容的日期解析：支持 Date、Excel 序列号、常见中英文日期格式 */
    function toDate(v) {
      if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
      if (v === null || v === undefined) return null;
      if (typeof v === 'number') {
        if (v > 59 && v < 2958466 && window.XLSX && XLSX.SSF && XLSX.SSF.parse_date_code) {
          const d = XLSX.SSF.parse_date_code(v);
          if (d) return new Date(d.y, (d.m || 1) - 1, d.d || 1, d.H || 0, d.M || 0, Math.floor(d.S || 0));
        }
        return null;
      }
      if (typeof v !== 'string') return null;
      const s = v.trim();
      if (!s || s.length < 5) return null;
      let m = s.match(/^(\d{4})[-/.年](\d{1,2})(?:[-/.月](\d{1,2}))?日?(?:[ T]+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
      if (m) {
        const d = new Date(+m[1], +m[2] - 1, +(m[3] || 1), +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
        return isNaN(d.getTime()) ? null : d;
      }
      m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
      if (m) { const d = new Date(+m[3], +m[1] - 1, +m[2]); return isNaN(d.getTime()) ? null : d; }
      if (/[a-z]/i.test(s) || (s.indexOf(':') > 0 && s.length >= 8)) {
        const t = Date.parse(s);
        if (!isNaN(t)) return new Date(t);
      }
      return null;
    }

    function cellKind(v) {
      if (isBlank(v)) return 'blank';
      if (v instanceof Date) return 'date';
      if (typeof v === 'boolean') return 'bool';
      if (typeof v === 'number') return 'num';
      const s = String(v).trim();
      if (/^(true|false|yes|no|y|n|是|否|对|错)$/i.test(s)) return 'bool';
      if (looksLikeDate(s)) return toDate(s) ? 'date' : 'text';
      return isNaN(toNum(s)) ? 'text' : 'num';
    }

    /** 根据抽样推断列类型：number / date / text / bool / empty */
    function inferType(values) {
      const n = values.length;
      if (!n) return 'empty';
      const step = Math.max(1, Math.floor(n / 600));
      const cnt = { num: 0, date: 0, text: 0, bool: 0, blank: 0 };
      let total = 0;
      for (let i = 0; i < n; i += step) {
        const k = cellKind(values[i]);
        cnt[k] = (cnt[k] || 0) + 1;
        if (k !== 'blank') total++;
      }
      if (!total) return 'empty';
      const cand = [['num', cnt.num], ['date', cnt.date], ['bool', cnt.bool], ['text', cnt.text]]
        .sort((a, b) => b[1] - a[1]);
      if (cand[0][1] / total >= 0.75) {
        if (cand[0][0] === 'num') return 'number';
        if (cand[0][0] === 'date') return 'date';
        if (cand[0][0] === 'bool') return 'bool';
        return 'text';
      }
      return 'text';
    }

    function mean(a) { if (!a.length) return 0; let s = 0; for (let i = 0; i < a.length; i++) s += a[i]; return s / a.length; }
    function quantile(sorted, p) {
      if (!sorted.length) return NaN;
      if (sorted.length === 1) return sorted[0];
      const pos = (sorted.length - 1) * p, lo = Math.floor(pos), hi = Math.ceil(pos);
      return lo === hi ? sorted[lo] : sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
    }
    function median(a) { const s = a.slice().sort((x, y) => x - y); return quantile(s, 0.5); }
    function stdev(a) {
      if (a.length < 2) return 0;
      const m = mean(a); let s = 0;
      for (let i = 0; i < a.length; i++) { const d = a[i] - m; s += d * d; }
      return Math.sqrt(s / (a.length - 1));
    }
    function pearson(a, b) {
      const n = Math.min(a.length, b.length);
      if (n < 2) return 0;
      const ma = mean(a), mb = mean(b);
      let num = 0, da = 0, db = 0;
      for (let i = 0; i < n; i++) {
        const x = a[i] - ma, y = b[i] - mb;
        num += x * y; da += x * x; db += y * y;
      }
      if (!da || !db) return 0;
      return num / Math.sqrt(da * db);
    }
    /** Sturges 公式自动分箱 + 等宽直方图 */
    function histogram(values, bins) {
      if (!values.length) return { bins: [], counts: [], min: 0, max: 0, width: 0 };
      const min = values.reduce((a, b) => a < b ? a : b), max = values.reduce((a, b) => a > b ? a : b);
      if (min === max) return { bins: [min], counts: [values.length], min, max, width: 0 };
      let k = bins > 0 ? bins : Math.max(4, Math.ceil(Math.log2(values.length) + 1));
      k = clamp(k, 2, 80);
      const w = (max - min) / k, counts = new Array(k).fill(0), edges = [];
      for (let i = 0; i <= k; i++) edges.push(min + w * i);
      for (let i = 0; i < values.length; i++) {
        let b = Math.floor((values[i] - min) / w);
        if (b >= k) b = k - 1; if (b < 0) b = 0;
        counts[b]++;
      }
      return { bins: edges, counts, min, max, width: w };
    }

    function fmtNum(v, digits) {
      if (v === null || v === undefined || v === '' || (typeof v === 'number' && !isFinite(v))) return '';
      const n = typeof v === 'number' ? v : toNum(v);
      if (isNaN(n)) return String(v);
      const d = digits === undefined ? (Math.abs(n) >= 1000 || Number.isInteger(n) ? 0 : 2) : digits;
      return n.toLocaleString('zh-CN', { maximumFractionDigits: d, minimumFractionDigits: 0 });
    }
    function fmtCompact(v) {
      const n = Number(v); if (isNaN(n)) return String(v);
      const a = Math.abs(n);
      if (a >= 1e8) return (n / 1e8).toFixed(2) + ' 亿';
      if (a >= 1e4) return (n / 1e4).toFixed(2) + ' 万';
      return fmtNum(n, a < 10 ? 2 : (a < 1000 ? 1 : 0));
    }
    function fmtDate(v, withTime) {
      const d = toDate(v); if (!d) return '';
      const p = n => String(n).padStart(2, '0');
      let s = d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
      if (withTime) s += ' ' + p(d.getHours()) + ':' + p(d.getMinutes()) + (d.getSeconds() ? ':' + p(d.getSeconds()) : '');
      return s;
    }
    function fmtBytes(n) {
      if (!n && n !== 0) return '—';
      const u = ['B', 'KB', 'MB', 'GB']; let i = 0, v = n;
      while (v >= 1024 && i < u.length - 1) { v /= 1024; i++; }
      return (i === 0 ? v : v.toFixed(1)) + ' ' + u[i];
    }
    function fmtVal(v, type) {
      if (isBlank(v)) return '';
      if (v instanceof Date) return fmtDate(v, !!(v.getHours() || v.getMinutes() || v.getSeconds()));
      if (type === 'number') { const n = toNum(v); return isNaN(n) ? String(v) : fmtNum(n); }
      if (type === 'date') { const d = toDate(v); return d ? fmtDate(d, !!(d.getHours() || d.getMinutes() || d.getSeconds())) : String(v); }
      if (type === 'bool') return /^(true|1|是|yes|y|对)$/i.test(String(v).trim()) ? '是' : '否';
      return String(v);
    }
    /** 单元格原始值 -> 可比较的排序键 */
    function sortKey(v, type) {
      if (isBlank(v)) return null;
      if (type === 'number') { const n = toNum(v); return isNaN(n) ? null : n; }
      if (type === 'date') { const d = toDate(v); return d ? d.getTime() : null; }
      return String(v);
    }

    function debounce(fn, ms) { let t; return function () { const a = arguments, c = this; clearTimeout(t); t = setTimeout(() => fn.apply(c, a), ms || 200); }; }
    function throttle(fn, ms) {
      let last = 0, timer = null;
      return function () {
        const now = Date.now(), a = arguments, c = this;
        const rest = (ms || 100) - (now - last);
        if (rest <= 0) { last = now; fn.apply(c, a); }
        else if (!timer) timer = setTimeout(() => { timer = null; last = Date.now(); fn.apply(c, a); }, rest);
      };
    }

    /** 通用分隔文本解析（CSV/TSV），支持引号、转义引号、换行内嵌 */
    function parseDelimited(text, delim) {
      const rows = []; let row = [], field = '', inQ = false;
      const d = delim || (guessDelim(text));
      const s = text.replace(/^\uFEFF/, '');
      for (let i = 0; i < s.length; i++) {
        const ch = s[i];
        if (inQ) {
          if (ch === '"') { if (s[i + 1] === '"') { field += '"'; i++; } else inQ = false; }
          else field += ch;
        } else if (ch === '"') inQ = true;
        else if (ch === d) { row.push(field); field = ''; }
        else if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
        else if (ch !== '\r') field += ch;
      }
      if (field.length || row.length) { row.push(field); rows.push(row); }
      while (rows.length && rows[rows.length - 1].every(c => String(c).trim() === '')) rows.pop();
      return rows;
    }
    function guessDelim(text) {
      const head = text.slice(0, 8192).split('\n').slice(0, 5).join('\n');
      const c = { '\t': (head.match(/\t/g) || []).length, ',': (head.match(/,/g) || []).length, ';': (head.match(/;/g) || []).length, '|': (head.match(/\|/g) || []).length };
      let best = ',', bv = -1;
      for (const k in c) if (c[k] > bv) { bv = c[k]; best = k; }
      return bv > 0 ? best : ',';
    }

    function toCSV(cols, rows) {
      const q = v => {
        if (v === null || v === undefined) return '';
        if (v instanceof Date) return '"' + fmtDate(v, true) + '"';
        const s = String(v);
        return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
      };
      const out = [cols.map(q).join(',')];
      for (let i = 0; i < rows.length; i++) out.push(rows[i].map(q).join(','));
      return out.join('\r\n');
    }
    function toTSV(cols, rows) {
      const q = v => v === null || v === undefined ? '' : String(v instanceof Date ? fmtDate(v, true) : v).replace(/[\t\n\r]/g, ' ');
      return [cols.map(q).join('\t')].concat(rows.map(r => r.map(q).join('\t'))).join('\n');
    }
    function downloadBlob(blob, filename) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = filename; document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 800);
    }
    function downloadText(text, filename, mime) {
      downloadBlob(new Blob([text], { type: (mime || 'text/plain') + ';charset=utf-8' }), filename);
    }
    async function copyText(text) {
      try {
        if (navigator.clipboard && window.isSecureContext) { await navigator.clipboard.writeText(text); return true; }
      } catch (e) { }
      try {
        const ta = document.createElement('textarea');
        ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
        document.body.appendChild(ta); ta.select();
        const ok = document.execCommand('copy'); ta.remove(); return ok;
      } catch (e) { return false; }
    }
    function stampName(base, ext) {
      const d = new Date(), p = n => String(n).padStart(2, '0');
      const s = d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '-' + p(d.getHours()) + p(d.getMinutes());
      return base.replace(/\.[^.]+$/, '') + '_' + s + '.' + ext;
    }
    /** 在转义后的文本中高亮关键词 */
    function hlText(text, tokens) {
      let out = esc(text);
      if (!tokens || !tokens.length) return out;
      for (const t of tokens) {
        if (!t) continue;
        const et = esc(t).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        out = out.replace(new RegExp(et, 'gi'), m => '\u0001' + m + '\u0002');
      }
      return out.replace(/\u0001/g, '<mark class="hl">').replace(/\u0002/g, '</mark>');
    }
    /* ==UTILS-END== */

/* =========================================================
       全局状态
       ========================================================= */
    const PALETTES = [
      { name: '靛蓝', c: ['#4f46e5', '#06b6d4', '#8b5cf6', '#f59e0b', '#10b981', '#ef4444', '#0ea5e9', '#ec4899', '#84cc16', '#f43f5e'] },
      { name: '海洋', c: ['#0ea5e9', '#22d3ee', '#3b82f6', '#6366f1', '#14b8a6', '#60a5fa', '#38bdf8', '#818cf8', '#2dd4bf', '#a5b4fc'] },
      { name: '暖阳', c: ['#f97316', '#f59e0b', '#ef4444', '#ec4899', '#f43f5e', '#fb923c', '#fbbf24', '#e11d48', '#facc15', '#fb7185'] },
      { name: '森林', c: ['#10b981', '#22c55e', '#84cc16', '#14b8a6', '#65a30d', '#059669', '#4ade80', '#0d9488', '#a3e635', '#34d399'] },
      { name: '莫兰迪', c: ['#8d99ae', '#b5838d', '#e5989b', '#6d6875', '#ffb4a2', '#a5a58d', '#b7b7a4', '#ffcb77', '#9a8c98', '#c9ada7'] },
      { name: '暗夜', c: ['#7c8cff', '#22d3ee', '#a78bfa', '#f472b6', '#34d399', '#fbbf24', '#60a5fa', '#fb7185', '#4ade80', '#c084fc'] }
    ];

    const CHART_TYPES = [
      { id: 'bar', n: '柱状图', ic: 'i-chart' }, { id: 'barh', n: '条形图', ic: 'i-rows' },
      { id: 'line', n: '折线图', ic: 'i-activity' }, { id: 'area', n: '面积图', ic: 'i-layers' },
      { id: 'pie', n: '饼图', ic: 'i-pie' }, { id: 'ring', n: '环形图', ic: 'i-target' },
      { id: 'scatter', n: '散点图', ic: 'i-dots' }, { id: 'radar', n: '雷达图', ic: 'i-star' },
      { id: 'funnel', n: '漏斗图', ic: 'i-filter' }, { id: 'box', n: '箱线图', ic: 'i-box' },
      { id: 'hist', n: '直方图', ic: 'i-hist' }, { id: 'heat', n: '相关性', ic: 'i-grid' }
    ];

    const AGG_NAMES = { sum: '求和', avg: '平均值', count: '计数', countAll: '行数', min: '最小值', max: '最大值', median: '中位数', distinct: '去重计数' };
    const TYPE_LABEL = { number: '数值', text: '文本', date: '日期', bool: '布尔', empty: '空' };
    const TYPE_CLS = { number: 't-num', text: 't-text', date: 't-date', bool: 't-bool', empty: 't-empty' };
    const TYPE_IC = { number: 'i-hash', text: 'i-text', date: 'i-calendar', bool: 'i-check', empty: 'i-minus' };

    const I18N = {
      '数据接入':'Data Import','数据概览':'Overview','数据表':'Data Table','可视化':'Visualization','透视分析':'Pivot Analysis','数据清洗':'Data Cleaning','报告导出':'Report Export',
      '开始':'START','分析':'ANALYSIS','加工':'PROCESS','数据源':'DATA SOURCE','未载入':'Not loaded','隐私优先':'Privacy first','本地处理':'Local processing','示例数据':'Demo data','导入数据':'Import data','命令':'Commands','帮助与快捷键':'Help & shortcuts','设置':'Settings','切换主题':'Toggle theme',
      '数据质量提醒':'Data quality alerts','列画像':'Column profiles','数值列速览':'Numeric columns','重新分析':'Reprofile','去可视化':'Open visualization','图表配置':'Chart settings','智能推荐':'Smart recommend','重置配置':'Reset settings','透视结果':'Pivot result','生成透视表':'Generate pivot table','操作历史':'Operation history','清洗前后对比':'Before / after','报告设置':'Report settings','预览':'Preview',
      '导出':'Export','刷新预览':'Refresh preview','打印 / 存为 PDF':'Print / Save PDF','下载 HTML 报告':'Download HTML report','显示列':'Columns','紧凑':'Compact','标准':'Standard','宽松':'Cozy','添加筛选':'Add filter','删除所选':'Delete selected','每页':'Rows per page','跳转到':'Go to','行':'rows',
      '图表类型':'Chart type','数据模式':'Data mode','按维度聚合':'Aggregate by dimension','逐行明细':'Raw rows','维度 / X 轴':'Dimension / X axis','系列拆分（可选）':'Split series (optional)','指标 / Y 轴':'Metrics / Y axis','聚合方式':'Aggregation','显示选项':'Display options','数据来源':'Data source','当前筛选结果':'Filtered rows','全部数据':'All rows',
      '行操作':'Row operations','缺失值处理':'Missing values','列操作':'Column operations','新增计算列':'New calculated column','撤销':'Undo','恢复原始数据':'Restore original','应用到该列':'Apply to column','删除列':'Delete column','重命名':'Rename','拆分':'Split','去重':'Deduplicate','预览前 5 行':'Preview first 5 rows','添加该列':'Add column','语法说明':'Syntax help','关闭':'Close','取消':'Cancel','确定':'Confirm','好的':'Got it',
      '求和 Sum':'Sum','平均值 Avg':'Average','计数 Count':'Count','最小值 Min':'Minimum','最大值 Max':'Maximum','中位数 Median':'Median','按总计降序排列':'Sort totals descending','显示总计行 / 列':'Show totals','按数值热力着色':'Heatmap values','显示占总计百分比':'Show percentage of total',
      '支持 100 万+ 单元格':'Supports 1M+ cells','版本 2.1 · 支持 CSV / TSV / XLSX / XLS / JSON · 最大建议 50 MB':'v2.1 · CSV / TSV / XLSX / XLS / JSON · 50 MB recommended max'
    };
    const I18N_REV = Object.fromEntries(Object.entries(I18N).map(([zh, en]) => [en, zh]));
    let i18nObserver = null;
    function applyLanguage(lang, quiet) {
      const next = lang === 'en' ? 'en' : 'zh';
      S.set.lang = next;
      document.documentElement.lang = next === 'en' ? 'en' : 'zh-CN';
      if (i18nObserver) i18nObserver.disconnect();
      const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
      const nodes = []; let node;
      while ((node = walker.nextNode())) nodes.push(node);
      nodes.forEach(n => { const raw = n.nodeValue, key = raw.trim(); if (!key) return; const value = next === 'en' ? I18N[key] : I18N_REV[key]; if (value) n.nodeValue = raw.replace(key, value); });
      $$('[title],[placeholder],[aria-label],[data-tip]').forEach(el => ['title','placeholder','aria-label','data-tip'].forEach(attr => { const key = el.getAttribute(attr); const value = next === 'en' ? I18N[key] : I18N_REV[key]; if (value) el.setAttribute(attr, value); }));
      if (D.langToggle) { D.langToggle.setAttribute('aria-label', next === 'en' ? '切换到中文' : 'Switch to English'); D.langToggle.setAttribute('data-tip', next === 'en' ? 'Switch to 中文' : '切换中文 / English'); D.langToggle.classList.toggle('is-en', next === 'en'); }
      saveSettings();
      i18nObserver = new MutationObserver(() => applyLanguage(S.set.lang, true));
      i18nObserver.observe(document.body, { childList: true, subtree: true });
      if (!quiet) toast(next === 'en' ? 'English mode enabled' : '已切换为中文模式', { kind: 'ok', title: next === 'en' ? 'Language' : '语言' });
    }
    function toggleLanguage() { applyLanguage(S.set.lang === 'en' ? 'zh' : 'en'); }

    const S = {
      cols: [], rows: [], origin: null,
      file: { name: '', size: 0, sheets: [], sheet: '', at: 0, source: '' },
      profile: null, issues: [],
      view: { search: '', tokens: [], sort: { c: -1, d: 0 }, filters: {}, hidden: [], page: 1, size: 25, density: 'normal', idx: [], stickyFirst: false },
      sel: new Set(),
      hist: [], future: [],
      saved: [],
      cfg: { type: 'bar', mode: 'agg', x: -1, y: [], split: -1, agg: 'sum', topN: 20, sortBy: 'xasc', bin: 14, smooth: true, stack: false, label: false, rotate: true, zoom: false, title: '', palette: 0 },
      pv: { rows: [], col: -1, val: -1, agg: 'sum', totals: true, heat: true, pct: false, sortDesc: true, scope: 'filtered', last: null },
      set: { theme: 'system', lang: 'zh', pageSize: 25, palette: 0, autosave: false, density: 'normal', seenIntro: false },
      chart: null, chartOpt: null, full: false,
      searchCache: null, dirty: false, lastSheet: null
    };

    const D = {};
    const IDS = ['topProgress', 'navToggle', 'topFileInfo', 'topFileName', 'topFileMeta', 'btnPalette', 'langToggle', 'btnDemo', 'btnOpen',
      'btnHelp', 'btnSettings', 'btnTheme', 'themeIcon', 'sidebar', 'navColCount', 'navRowCount', 'navHistCount',
      'srcBadge', 'srcDetail', 'srcMeter', 'srcComplete', 'dropzone', 'btnPaste', 'btnDemo2', 'btnRestore', 'optAutosave',
      'btnClearRecent', 'recentList', 'ovSubtitle', 'btnReprofile', 'btnOvToChart', 'kpiGrid', 'profCount', 'profSort',
      'profBody', 'issueCount', 'issueList', 'numQuick', 'tblSubtitle', 'btnDelSel', 'selCount', 'btnCols', 'tblDensity',
      'btnExportTbl', 'tblSearch', 'btnSearchClear', 'btnAddFilter', 'filterChips', 'tblCount', 'tblEmpty', 'tblWrap', 'tbl',
      'tblHead', 'tblBody', 'tblRange', 'pageSize', 'pageJump', 'pager', 'chartSubtitle', 'btnChartAuto', 'btnChartReset',
      'chartTypeGrid', 'modeBadge', 'chartMode', 'modeHint', 'cfgX', 'wrapSplit', 'cfgSplit', 'yList', 'yHint', 'wrapAgg',
      'cfgAgg', 'wrapTopN', 'cfgTopN', 'wrapSortBy', 'cfgSortBy', 'wrapBin', 'cfgBin', 'cfgBinVal', 'cfgSmooth', 'cfgStack',
      'cfgLabel', 'cfgRotate', 'cfgDataZoom', 'paletteList', 'cfgTitle', 'btnRenderChart', 'btnSaveChart', 'chartTitle',
      'btnChartFull', 'btnChartPng', 'btnChartJson', 'chartBox', 'chartEmpty', 'chartFoot', 'savedCount', 'btnClearSaved',
      'savedGrid', 'btnPivotCopy', 'btnPivotExport', 'pvRowList', 'pvCol', 'pvVal', 'pvAgg', 'pvTotals', 'pvHeat',
      'pvPercent', 'pvSortDesc', 'pvScope', 'btnPivotRun', 'pvTitle', 'pvInfo', 'pvEmpty', 'pvWrap', 'btnUndo',
      'btnResetData', 'cleanRowInfo', 'fillCol', 'fillMode', 'fillText', 'btnFill', 'btnFillAllNumeric', 'fillPreview',
      'colOp', 'btnRenameCol', 'btnTypeNum', 'btnTypeDate', 'btnTypeText', 'btnDropCol', 'splitCol', 'splitSep', 'btnSplit',
      'dedupeCol', 'btnDedupe', 'calcName', 'calcExpr', 'btnCalcPreview', 'btnCalcApply', 'btnCalcHelp', 'calcPreview',
      'calcQuick', 'histCount', 'histList', 'cleanDiff', 'btnReportRefresh', 'btnPrint', 'btnExportHtml', 'rpTitle',
      'rpAuthor', 'rpSummary', 'rpIncKpi', 'rpIncProfile', 'rpIncIssues', 'rpIncCharts', 'rpIncPivot', 'rpIncSample',
      'rpSampleWrap', 'rpSampleRows', 'reportPreview', 'reportPrintArea', 'popRoot', 'modalRoot', 'drawerRoot', 'cmdkRoot',
      'toasts', 'fileInput', 'linkHelp', 'linkAbout'];

    function cacheDom() {
      IDS.forEach(id => {
        const e = document.getElementById(id);
        if (e) D[id] = e;
        else console.warn('[DataLens] 缺少元素 #' + id);
      });
    }

    /* =========================================================
       通用 UI：进度条 / Toast / 模态 / 气泡菜单 / 抽屉
       ========================================================= */
    function progress(on, indet) {
      const p = D.topProgress; if (!p) return;
      if (on) { p.classList.add('on'); p.classList.toggle('indet', indet !== false); }
      else { p.classList.remove('on', 'indet'); p.style.transform = 'scaleX(0)'; }
    }
    function nextFrame() { return new Promise(r => requestAnimationFrame(() => setTimeout(r, 0))); }

    const TOAST_IC = { ok: 'i-check', err: 'i-alert', warn: 'i-alert', info: 'i-info' };
    function toast(msg, opt) {
      opt = opt || {};
      const kind = opt.kind || 'info';
      const box = document.createElement('div');
      box.className = 'toast ' + kind;
      box.innerHTML =
        '<span class="t-ic"><svg class="ic ic-sm"><use href="#' + (TOAST_IC[kind] || 'i-info') + '"/></svg></span>' +
        '<span class="t-msg">' + (opt.title ? '<span class="t-title">' + esc(opt.title) + '</span>' : '') +
        '<span class="t-desc">' + esc(msg) + '</span></span>' +
        (opt.action ? '<button class="t-act">' + esc(opt.action) + '</button>' : '') +
        '<button class="t-x" aria-label="关闭"><svg class="ic ic-sm"><use href="#i-x"/></svg></button>';
      const close = () => { if (!box.isConnected) return; box.classList.add('out'); setTimeout(() => box.remove(), 240); };
      box.querySelector('.t-x').onclick = close;
      if (opt.action) box.querySelector('.t-act').onclick = () => { close(); opt.onAction && opt.onAction(); };
      D.toasts.appendChild(box);
      const life = opt.sticky ? 0 : (opt.ms || (kind === 'err' ? 6500 : 3600));
      if (life) setTimeout(close, life);
      while (D.toasts.children.length > 5) D.toasts.firstElementChild.remove();
      return close;
    }

    let modalStack = [];
    function openModal(opt) {
      const mask = document.createElement('div');
      mask.className = 'mask';
      const m = document.createElement('div');
      m.className = 'modal ' + (opt.size || '');
      m.setAttribute('role', 'dialog'); m.setAttribute('aria-modal', 'true');
      m.innerHTML =
        (opt.title === null ? '' : '<div class="modal-h"><h3>' + (opt.icon ? '<svg class="ic" style="color:var(--brand)"><use href="#' + opt.icon + '"/></svg>' : '') + esc(opt.title || '') + '</h3><div class="grow"></div><button class="btn btn-ghost btn-icon btn-sm" data-close aria-label="关闭"><svg class="ic ic-sm"><use href="#i-x"/></svg></button></div>') +
        '<div class="modal-b">' + (opt.body || '') + '</div>' +
        (opt.footer === null ? '' : '<div class="modal-f">' + (opt.footer || '<button class="btn" data-close>关闭</button>') + '</div>');
      mask.appendChild(m);
      D.modalRoot.appendChild(mask);
      const api = {
        root: m, mask,
        close() {
          mask.style.animation = 'fade .16s reverse both';
          setTimeout(() => { mask.remove(); modalStack = modalStack.filter(x => x !== api); }, 150);
          opt.onClose && opt.onClose();
        }
      };
      modalStack.push(api);
      mask.addEventListener('mousedown', e => { if (e.target === mask && opt.dismissable !== false) api.close(); });
      m.querySelectorAll('[data-close]').forEach(b => b.onclick = () => api.close());
      setTimeout(() => {
        const f = m.querySelector('[data-autofocus]') || m.querySelector('input,textarea,select,button');
        f && f.focus();
      }, 40);
      opt.onMount && opt.onMount(m, api);
      return api;
    }
    function topModal() { return modalStack[modalStack.length - 1]; }

    function confirmDlg(opt) {
      return new Promise(resolve => {
        let done = false;
        const finish = v => { if (done) return; done = true; resolve(v); };
        const m = openModal({
          title: opt.title || '请确认', icon: opt.icon || 'i-alert', size: 'slim',
          body: '<p class="muted" style="font-size:13.5px">' + esc(opt.message || '') + '</p>' +
            (opt.detail ? '<p class="tiny dim" style="margin-top:8px">' + esc(opt.detail) + '</p>' : ''),
          footer: '<button class="btn" data-no>' + esc(opt.cancelText || '取消') + '</button>' +
            '<button class="btn ' + (opt.danger ? 'btn-danger' : 'btn-primary') + '" data-yes data-autofocus>' + esc(opt.okText || '确定') + '</button>',
          onClose: () => finish(false)
        });
        m.root.querySelector('[data-no]').onclick = () => { finish(false); m.close(); };
        m.root.querySelector('[data-yes]').onclick = () => { finish(true); m.close(); };
      });
    }

    function promptDlg(opt) {
      return new Promise(resolve => {
        let done = false;
        const finish = v => { if (done) return; done = true; resolve(v); };
        const m = openModal({
          title: opt.title || '请输入', icon: opt.icon || 'i-pencil', size: 'slim',
          body: '<div class="field"><label>' + esc(opt.label || '') + '</label>' +
            '<input class="input" id="pv-inp" value="' + esc(opt.value || '') + '" placeholder="' + esc(opt.placeholder || '') + '" data-autofocus>' +
            (opt.hint ? '<span class="hint">' + esc(opt.hint) + '</span>' : '') + '</div>',
          footer: '<button class="btn" data-no>取消</button><button class="btn btn-primary" data-yes>确定</button>',
          onClose: () => finish(null)
        });
        const inp = m.root.querySelector('#pv-inp');
        const yes = () => { const v = inp.value; if (opt.required && !v.trim()) { inp.focus(); toast('内容不能为空', { kind: 'warn' }); return; } finish(v); m.close(); };
        m.root.querySelector('[data-no]').onclick = () => { finish(null); m.close(); };
        m.root.querySelector('[data-yes]').onclick = yes;
        inp.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); yes(); } };
        setTimeout(() => inp.select(), 60);
      });
    }

    let curPop = null;
    function closePop() { if (curPop) { curPop.remove(); curPop = null; } }
    function openPop(anchor, html, opt) {
      opt = opt || {};
      closePop();
      const p = document.createElement('div');
      p.className = 'pop';
      p.innerHTML = html;
      document.body.appendChild(p);
      const r = anchor.getBoundingClientRect();
      const pw = p.offsetWidth, ph = p.offsetHeight;
      let left = opt.align === 'right' ? r.right - pw : r.left;
      left = clamp(left, 8, window.innerWidth - pw - 8);
      let top = r.bottom + 6;
      if (top + ph > window.innerHeight - 8) top = Math.max(8, r.top - ph - 6);
      p.style.left = left + 'px'; p.style.top = top + 'px';
      curPop = p;
      const away = e => {
        const t = e && e.target;
        if (t && t.nodeType === 1) {
          if (!p.contains(t) && !(anchor && anchor.nodeType === 1 && anchor.contains(t))) { closePop(); cleanup(); }
          return;
        }
        closePop(); cleanup(); // resize / 窗口滚动等非元素事件：直接收起
      };
      const key = e => { if (e.key === 'Escape') { closePop(); cleanup(); } };
      const cleanup = () => { document.removeEventListener('mousedown', away, true); document.removeEventListener('keydown', key, true); window.removeEventListener('resize', away); window.removeEventListener('scroll', away, true); };
      setTimeout(() => {
        document.addEventListener('mousedown', away, true);
        document.addEventListener('keydown', key, true);
        window.addEventListener('resize', away);
        window.addEventListener('scroll', away, true);
      }, 0);
      opt.onMount && opt.onMount(p, () => { closePop(); cleanup(); });
      return p;
    }

    function openDrawer(title, bodyHtml, opt) {
      opt = opt || {};
      D.drawerRoot.innerHTML =
        '<div class="drawer-mask"></div><div class="drawer" role="dialog" aria-modal="true" aria-label="' + esc(title) + '">' +
        '<div class="drawer-h"><svg class="ic" style="color:var(--brand)"><use href="#' + (opt.icon || 'i-rows') + '"/></svg>' +
        '<h3 class="b grow truncate" style="font-size:15px">' + esc(title) + '</h3>' +
        (opt.actions || '') +
        '<button class="btn btn-ghost btn-icon btn-sm" data-dclose aria-label="关闭"><svg class="ic ic-sm"><use href="#i-x"/></svg></button></div>' +
        '<div class="drawer-b">' + bodyHtml + '</div></div>';
      const close = () => { D.drawerRoot.innerHTML = ''; };
      D.drawerRoot.querySelector('.drawer-mask').onclick = close;
      D.drawerRoot.querySelectorAll('[data-dclose]').forEach(b => b.onclick = close);
      opt.onMount && opt.onMount(D.drawerRoot.querySelector('.drawer'), close);
      return { close, root: D.drawerRoot.querySelector('.drawer') };
    }

/* =========================================================
       数据接入：解析文件 / 文本 / 示例数据
       ========================================================= */
    function buildCols(headerRow, nCols) {
      const used = Object.create(null), cols = [];
      for (let i = 0; i < nCols; i++) {
        let raw = headerRow && headerRow[i] !== undefined && headerRow[i] !== null ? String(headerRow[i]).trim() : '';
        raw = raw.replace(/\s+/g, ' ').slice(0, 60);
        if (!raw) raw = '列' + (i + 1);
        let name = raw, k = 2;
        while (used[name.toLowerCase()]) { name = raw + '_' + k; k++; }
        used[name.toLowerCase()] = 1;
        cols.push({ name, type: 'text' });
      }
      return cols;
    }

    /** 把二维数组装载为数据集 */
    function ingestMatrix(matrix, meta) {
      if (!matrix || !matrix.length) { toast('文件里没有可读取的数据', { kind: 'err', title: '导入失败' }); return false; }
      let maxLen = 0;
      for (let i = 0; i < Math.min(matrix.length, 2000); i++) maxLen = Math.max(maxLen, (matrix[i] || []).length);
      for (let i = 0; i < matrix.length; i++) maxLen = Math.max(maxLen, (matrix[i] || []).length);
      if (!maxLen) { toast('未识别到任何列', { kind: 'err' }); return false; }

      const first = (matrix[0] || []).slice(0, maxLen);
      const nonEmpty = first.filter(v => !isBlank(v));
      const numericFirst = nonEmpty.length > 0 && nonEmpty.every(v => cellKind(v) === 'num');
      const useHeader = meta && meta.forceNoHeader ? false : !numericFirst;

      const cols = useHeader ? buildCols(first, maxLen) : buildCols(null, maxLen);
      const startRow = useHeader ? 1 : 0;
      const rows = [];
      for (let r = startRow; r < matrix.length; r++) {
        const src = matrix[r] || [];
        const out = new Array(maxLen);
        for (let c = 0; c < maxLen; c++) out[c] = src[c] === undefined ? '' : src[c];
        if (out.every(isBlank)) continue; // 跳过整行空白
        rows.push(out);
      }
      if (!rows.length) { toast('除表头外没有数据行', { kind: 'warn', title: '导入未完成' }); return false; }

      S.cols = cols; S.rows = rows;
      S.origin = { cols: cols.map(c => ({ ...c })), rows: rows.map(r => r.slice()) };
      S.hist = []; S.future = []; S.sel.clear();
      S.view.filters = {}; S.view.search = ''; S.view.tokens = []; S.view.sort = { c: -1, d: 0 };
      S.view.page = 1; S.view.hidden = []; S.searchCache = null;
      S.saved = []; S.pv.last = null; S.pv.rows = []; S.pv.col = -1; S.pv.val = -1;
      S.dirty = false;

      cols.forEach((c, i) => { c.type = inferType(rows.map(r => r[i])); });

      S.file = Object.assign({ name: '未命名数据', size: 0, sheets: [], sheet: '', at: Date.now(), source: '导入' }, meta || {});
      if (rows.length > 200000) toast('数据量较大（' + fmtNum(rows.length) + ' 行），部分操作可能需要几秒', { kind: 'warn', title: '性能提示', ms: 6000 });
      if (!useHeader) toast('首行看起来是数据，已自动生成列名（列1、列2…）', { kind: 'info', title: '未检测到表头', action: '去重命名', onAction: () => { go('clean'); } });
      return true;
    }

    function loadSheetFromWorkbook(wb, name) {
      const ws = wb.Sheets[name];
      if (!ws) return false;
      const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: false });
      return ingestMatrix(matrix, { sheet: name });
    }

    async function handleFiles(fileList) {
      const files = Array.from(fileList || []);
      if (!files.length) return;
      const file = files[0];
      if (files.length > 1) toast('已忽略其余 ' + (files.length - 1) + ' 个文件，一次仅分析一个数据集', { kind: 'info' });
      if (file.size > 80 * 1024 * 1024) {
        const ok = await confirmDlg({ title: '文件较大', message: '该文件约 ' + fmtBytes(file.size) + '，浏览器解析可能较慢甚至卡住。', detail: '建议先拆分为更小的文件，或继续尝试。', okText: '继续尝试', danger: true });
        if (!ok) return;
      }
      progress(true);
      await nextFrame();
      const ext = (file.name.split('.').pop() || '').toLowerCase();
      try {
        if (ext === 'json') {
          const text = await file.text();
          loadJsonText(text, { name: file.name, size: file.size, source: '文件' });
        } else {
          if (!window.XLSX) throw new Error('表格解析库（SheetJS）未能加载，请检查网络后刷新页面');
          const buf = await file.arrayBuffer();
          const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true, dense: false });
          S.lastSheet = { wb, name: file.name, size: file.size };
          const names = wb.SheetNames || [];
          if (!names.length) throw new Error('工作簿中没有工作表');
          if (!loadSheetFromWorkbook(wb, names[0])) return;
          S.file.sheets = names;
          S.file.name = file.name; S.file.size = file.size; S.file.source = '文件';
          if (names.length > 1) toast('检测到 ' + names.length + ' 个工作表，已载入「' + names[0] + '」', { kind: 'info', title: '多工作表', action: '切换', onAction: showSheetPicker });
        }
        afterLoad();
      } catch (err) {
        console.error(err);
        toast(String(err && err.message || err), { kind: 'err', title: '导入失败', ms: 8000 });
      } finally { progress(false); }
    }

    function loadJsonText(text, meta) {
      let data;
      try { data = JSON.parse(text); }
      catch (e) { throw new Error('JSON 解析失败：' + e.message); }
      if (data && !Array.isArray(data) && Array.isArray(data.data)) data = data.data;
      if (data && !Array.isArray(data) && Array.isArray(data.rows)) data = data.rows;
      let matrix;
      if (Array.isArray(data)) {
        if (!data.length) throw new Error('JSON 数组为空');
        if (Array.isArray(data[0])) matrix = data;
        else if (data[0] && typeof data[0] === 'object') {
          const keys = [];
          data.forEach(o => Object.keys(o || {}).forEach(k => { if (keys.indexOf(k) < 0) keys.push(k); }));
          matrix = [keys].concat(data.map(o => keys.map(k => {
            const v = (o || {})[k];
            return v === null || v === undefined ? '' : (typeof v === 'object' ? JSON.stringify(v) : v);
          })));
        } else matrix = [['值']].concat(data.map(v => [v]));
      } else throw new Error('暂不支持的 JSON 结构，请提供对象数组或二维数组');
      if (!ingestMatrix(matrix, meta)) return;
      S.file = Object.assign(S.file, meta || {});
    }

    function showSheetPicker() {
      if (!S.lastSheet) return;
      const names = S.lastSheet.wb.SheetNames || [];
      openModal({
        title: '选择工作表', icon: 'i-layers', size: 'slim',
        body: '<div class="col gap-2">' + names.map(n => {
          const ws = S.lastSheet.wb.Sheets[n];
          const ref = ws && ws['!ref'] ? ws['!ref'] : '空';
          return '<button class="pop-item" data-sheet="' + esc(n) + '" style="border:1px solid var(--border)">' +
            '<svg class="ic ic-sm" style="color:var(--brand)"><use href="#i-table"/></svg>' + esc(n) +
            '<span class="grow"></span><span class="tiny dim mono">' + esc(ref) + '</span></button>';
        }).join('') + '</div>',
        footer: '<button class="btn" data-close>取消</button>',
        onMount(root, api) {
          root.querySelectorAll('[data-sheet]').forEach(b => b.onclick = () => {
            const n = b.getAttribute('data-sheet');
            progress(true);
            setTimeout(() => {
              try {
                if (loadSheetFromWorkbook(S.lastSheet.wb, n)) {
                  S.file.sheets = names; S.file.name = S.lastSheet.name; S.file.size = S.lastSheet.size; S.file.source = '文件';
                  afterLoad(); toast('已切换到工作表「' + n + '」', { kind: 'ok' });
                }
              } catch (e) { toast(String(e.message || e), { kind: 'err' }); }
              finally { progress(false); api.close(); }
            }, 30);
          });
        }
      });
    }

    /** 生成示例数据：电商销售明细 */
    function loadDemo() {
      const regions = ['华东', '华北', '华南', '西南', '东北', '西北'];
      const cities = { '华东': ['上海', '杭州', '南京', '苏州'], '华北': ['北京', '天津', '石家庄'], '华南': ['广州', '深圳', '厦门'], '西南': ['成都', '重庆', '昆明'], '东北': ['沈阳', '大连', '哈尔滨'], '西北': ['西安', '兰州', '乌鲁木齐'] };
      const cats = ['数码', '家电', '服饰', '美妆', '食品', '家居', '运动'];
      const prods = { '数码': ['无线耳机', '机械键盘', '4K 显示器', '智能手表'], '家电': ['空气炸锅', '扫地机器人', '变频空调', '洗烘一体机'], '服饰': ['轻量羽绒服', '纯棉 T 恤', '休闲牛仔裤', '运动外套'], '美妆': ['精华液', '防晒霜', '口红套装', '面膜'], '食品': ['坚果礼盒', '精品咖啡', '有机茶叶', '进口零食'], '家居': ['四件套', '收纳柜', '香薰灯', '实木餐桌'], '运动': ['瑜伽垫', '跑步鞋', '哑铃套装', '登山包'] };
      const chans = ['天猫', '京东', '抖音', '线下门店', '微信小程序'];
      const reps = ['张伟', '李娜', '王强', '刘洋', '陈静', '赵磊', '孙悦', '周涛'];
      const rnd = mulberry(20260909);
      const head = ['订单日期', '区域', '城市', '品类', '商品', '渠道', '销售员', '数量', '单价', '折扣', '销售额', '成本', '利润', '满意度', '是否退货'];
      const N = 600, rows = [head];
      const start = new Date(2025, 0, 1).getTime(), end = new Date(2025, 11, 31).getTime();
      for (let i = 0; i < N; i++) {
        const region = regions[Math.floor(rnd() * regions.length)];
        const city = cities[region][Math.floor(rnd() * cities[region].length)];
        const cat = cats[Math.floor(rnd() * cats.length)];
        const prod = prods[cat][Math.floor(rnd() * prods[cat].length)];
        const qty = 1 + Math.floor(rnd() * 12);
        const price = Math.round((39 + rnd() * 2600) * 100) / 100;
        const disc = [0, 0, 0, 0.05, 0.1, 0.15, 0.2, 0.3][Math.floor(rnd() * 8)];
        const sales = Math.round(qty * price * (1 - disc) * 100) / 100;
        const cost = Math.round(sales * (0.45 + rnd() * 0.3) * 100) / 100;
        const dt = new Date(start + rnd() * (end - start));
        const d = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
        const ret = rnd() < 0.07;
        rows.push([fmtDate(d), region, city, cat, prod, chans[Math.floor(rnd() * chans.length)],
        reps[Math.floor(rnd() * reps.length)], qty, price, disc, sales, cost,
          Math.round((sales - cost) * 100) / 100, Math.round((3 + rnd() * 2) * 10) / 10, ret ? '是' : '否']);
      }
      // 注入少量缺失值，便于演示数据质量分析
      for (let i = 1; i < rows.length; i++) {
        if (rnd() < 0.05) rows[i][6] = '';
        if (rnd() < 0.03) rows[i][13] = '';
        if (rnd() < 0.02) rows[i][9] = '';
      }
      if (!ingestMatrix(rows, { name: '示例数据-2025销售明细.csv', size: 0, source: '示例数据' })) return;
      afterLoad();
      toast('已载入 600 行示例销售数据，可自由体验全部功能', { kind: 'ok', title: '示例数据就绪' });
    }
    function mulberry(seed) {
      let a = seed >>> 0;
      return function () {
        a |= 0; a = a + 0x6D2B79F5 | 0;
        let t = Math.imul(a ^ a >>> 15, 1 | a);
        t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
        return ((t ^ t >>> 14) >>> 0) / 4294967296;
      };
    }

    /* =========================================================
       数据画像
       ========================================================= */
    function profileCol(i) {
      const rows = S.rows, n = rows.length, type = S.cols[i].type;
      const p = { i, name: S.cols[i].name, type, total: n, blank: 0, unique: 0, missing: 0, top: [], dist: null, sample: null, avgLen: 0, unparse: 0 };
      const seen = new Set(), freq = new Map(), nums = [];
      let lenSum = 0, dmin = null, dmax = null;
      const storeSorted = n <= 200000;
      for (let r = 0; r < n; r++) {
        const v = rows[r][i];
        if (isBlank(v)) { p.blank++; continue; }
        if (p.sample === null) p.sample = v;
        const key = v instanceof Date ? 'd' + v.getTime() : (typeof v === 'number' ? 'n' + v : String(v));
        if (seen.size < 1200000) seen.add(key);
        if (type === 'number') {
          const x = toNum(v);
          if (isNaN(x)) p.unparse++; else nums.push(x);
        } else if (type === 'date') {
          const d = toDate(v);
          if (d) { const t = d.getTime(); if (dmin === null || t < dmin) dmin = t; if (dmax === null || t > dmax) dmax = t; }
          else p.unparse++;
        } else {
          const s = String(v); lenSum += s.length;
          if (freq.size < 6000) freq.set(s, (freq.get(s) || 0) + 1);
        }
      }
      p.unique = seen.size;
      p.missing = n ? p.blank / n : 0;
      const nonBlank = n - p.blank;
      if (type === 'number' && nums.length) {
        const a = nums.slice().sort((x, y) => x - y);
        p.count = a.length; p.min = a[0]; p.max = a[a.length - 1];
        p.mean = mean(a); p.median = quantile(a, .5); p.q1 = quantile(a, .25); p.q3 = quantile(a, .75);
        p.std = stdev(a); p.sum = a.reduce((s, x) => s + x, 0);
        if (storeSorted) p.sorted = a;
        p.dist = histogram(a, 16).counts;
        if (p.unique <= 14) {
          const f = new Map(); for (const x of a) f.set(x, (f.get(x) || 0) + 1);
          p.top = Array.from(f.entries()).sort((u, v) => v[1] - u[1]).slice(0, 3).map(e => [fmtNum(e[0]), e[1]]);
        }
        const iqr = p.q3 - p.q1;
        if (iqr > 0) {
          let out = 0; for (const x of a) if (x < p.q1 - 3 * iqr || x > p.q3 + 3 * iqr) out++;
          p.outliers = out;
        } else p.outliers = 0;
      } else if (type === 'date') {
        p.dmin = dmin; p.dmax = dmax;
        if (dmin !== null) p.spanDays = Math.round((dmax - dmin) / 86400000);
      } else {
        p.avgLen = nonBlank ? lenSum / nonBlank : 0;
        p.top = Array.from(freq.entries()).sort((a, b) => b[1] - a[1]).slice(0, 3);
        p.dist = p.top.map(t => t[1]);
      }
      return p;
    }

    function reprofile() {
      const t0 = performance.now();
      S.profile = S.cols.map((c, i) => profileCol(i));
      S.issues = detectIssues();
      return performance.now() - t0;
    }

    function detectIssues() {
      const out = [], P = S.profile || [];
      P.forEach(p => {
        if (p.type === 'empty') out.push({ lv: 'warn', ic: 'i-minus', t: '空列：' + p.name, d: '该列没有任何有效值，可以直接删除以简化数据结构。', act: { l: '删除该列', f: () => dropCol(p.i) } });
        else if (p.missing > 0.6) out.push({ lv: 'err', ic: 'i-alert', t: '严重缺失：' + p.name, d: '缺失率 ' + (p.missing * 100).toFixed(1) + '%（' + p.blank + ' 行），会显著影响统计结论。', act: { l: '去处理', f: () => gotoFill(p.i) } });
        else if (p.missing > 0.15) out.push({ lv: 'warn', ic: 'i-alert', t: '存在缺失：' + p.name, d: '缺失率 ' + (p.missing * 100).toFixed(1) + '%（' + p.blank + ' 行）。', act: { l: '去填充', f: () => gotoFill(p.i) } });
        if (p.unique <= 1 && p.total - p.blank > 1) out.push({ lv: 'info', ic: 'i-minus', t: '常量列：' + p.name, d: '所有非空值都相同（' + fmtVal(p.sample, p.type) + '），对分析没有区分度。', act: { l: '删除该列', f: () => dropCol(p.i) } });
        if (p.unparse > 0 && p.unparse / Math.max(1, p.total - p.blank) > 0.03) out.push({ lv: 'warn', ic: 'i-fx', t: '类型不一致：' + p.name, d: '识别为' + TYPE_LABEL[p.type] + '列，但有 ' + p.unparse + ' 个值无法转换，聚合时会被忽略。', act: { l: '去转换', f: () => { go('clean'); D.colOp.value = String(p.i); } } });
        if (p.type === 'number' && p.outliers > 0) out.push({ lv: 'info', ic: 'i-activity', t: '离群值：' + p.name, d: '按 3×IQR 规则发现 ' + p.outliers + ' 个极端值（范围 ' + fmtCompact(p.min) + ' ~ ' + fmtCompact(p.max) + '）。', act: { l: '看箱线图', f: () => { go('charts'); S.cfg.type = 'box'; S.cfg.x = -1; S.cfg.y = [p.i]; syncCfgUI(); renderChart(); } } });
      });
      const nonBlankTotal = P.reduce((s, p) => s + (p.total - p.blank), 0);
      const idCols = P.filter(p => p.unique === p.total - p.blank && p.total - p.blank > 20 && (p.type === 'number' || p.type === 'text'));
      if (idCols.length) out.push({ lv: 'info', ic: 'i-hash', t: '疑似主键列', d: idCols.map(p => p.name).join('、') + ' 每行唯一，适合作为记录标识，不适合做分组维度。' });
      const dup = countDupRows();
      if (dup > 0) out.push({ lv: 'warn', ic: 'i-copy', t: '重复行 ' + dup + ' 条', d: '完全相同的行有 ' + dup + ' 条多余副本，会让求和类指标虚高。', act: { l: '一键去重', f: () => runClean('dropDupRows') } });
      if (nonBlankTotal && P.some(p => p.missing > 0)) {
        const missCells = P.reduce((s, p) => s + p.blank, 0);
        out.push({ lv: missCells / (S.rows.length * Math.max(1, S.cols.length)) > 0.1 ? 'warn' : 'info', ic: 'i-target', t: '整体缺失 ' + fmtNum(missCells) + ' 个单元格', d: '占全部单元格的 ' + (missCells / (S.rows.length * Math.max(1, S.cols.length)) * 100).toFixed(1) + '%，可在「数据清洗」中批量填充。', act: { l: '去清洗', f: () => go('clean') } });
      }
      if (!out.length) out.push({ lv: 'ok', ic: 'i-check', t: '未发现明显质量问题', d: '列类型一致、缺失率较低、没有完全重复的行，数据可直接用于分析。' });
      const order = { err: 0, warn: 1, info: 2, ok: 3 };
      return out.sort((a, b) => order[a.lv] - order[b.lv]).slice(0, 24);
    }

    function countDupRows() {
      const seen = new Set(); let dup = 0;
      const rows = S.rows;
      for (let i = 0; i < rows.length; i++) {
        const k = rows[i].map(v => v instanceof Date ? v.getTime() : String(v)).join('\u0001');
        if (seen.has(k)) dup++; else seen.add(k);
      }
      return dup;
    }
    function gotoFill(i) { go('clean'); D.fillCol.value = String(i); D.fillCol.dispatchEvent(new Event('change')); }

/* =========================================================
       视图切换 / 侧边栏 / 数据源卡片
       ========================================================= */
    let currentView = 'ingest';
    function go(v) {
      if (!S.rows.length && ['overview', 'table', 'charts', 'pivot', 'clean', 'report'].indexOf(v) >= 0) v = 'ingest';
      currentView = v;
      $$('.nav-item').forEach(b => b.classList.toggle('active', b.dataset.view === v));
      $$('.view').forEach(s => s.classList.toggle('active', s.id === 'view-' + v));
      document.body.classList.remove('nav-open');
      window.scrollTo({ top: 0, behavior: 'smooth' });
      if (v === 'charts') setTimeout(() => { if (S.chart) S.chart.resize(); else renderChart(); }, 60);
      if (v === 'report') buildReport();
      if (v === 'table') setTimeout(() => D.tblSearch && D.tblSearch.focus(), 120);
      if (v === 'clean') renderCleanUI();
    }

    function enableNav(on) {
      $$('.nav-item.needs-data').forEach(b => { b.disabled = !on; });
    }

    function updateSourceCard() {
      const f = S.file;
      D.topFileInfo.classList.toggle('hidden', !S.rows.length);
      D.topFileName.textContent = f.name || '未命名数据';
      D.topFileMeta.textContent = S.rows.length ? fmtNum(S.rows.length) + ' 行 × ' + S.cols.length + ' 列' + (f.size ? ' · ' + fmtBytes(f.size) : '') : '';
      D.srcBadge.textContent = S.rows.length ? (f.sheet || f.source || '已载入') : '未载入';
      D.srcBadge.className = 'badge ' + (S.rows.length ? 'badge-ok' : '');
      const mix = S.profile ? countTypes(S.profile) : null;
      D.srcDetail.innerHTML = S.rows.length
        ? esc(f.name) + '<br>' + (f.sheets && f.sheets.length > 1 ? '<button class="btn btn-xs btn-soft" id="srcSheetBtn" style="margin:6px 0"><svg class="ic ic-sm"><use href="#i-layers"/></svg>切换工作表（' + f.sheets.length + '）</button><br>' : '')
        + '<span class="mono">' + fmtNum(S.rows.length) + '</span> 行 · <span class="mono">' + S.cols.length + '</span> 列'
        + (mix ? '<br>' + mix.map(m => '<span class="tag-type ' + TYPE_CLS[m.t] + '">' + TYPE_LABEL[m.t] + ' ' + m.n + '</span>').join(' ') : '')
        + '<br><span class="dim">' + new Date(f.at || Date.now()).toLocaleString('zh-CN') + '</span>'
        : '导入 CSV / Excel / JSON 后，这里会显示工作表、行列规模与类型分布。';
      const b = document.getElementById('srcSheetBtn'); if (b) b.onclick = showSheetPicker;
      const cells = S.rows.length * S.cols.length;
      const miss = S.profile ? S.profile.reduce((s, p) => s + p.blank, 0) : 0;
      const complete = cells ? (1 - miss / cells) : 0;
      D.srcMeter.style.width = (complete * 100).toFixed(1) + '%';
      D.srcComplete.textContent = cells ? (complete * 100).toFixed(1) + '%' : '—';
      D.navRowCount.textContent = fmtNum(S.rows.length);
      D.navColCount.textContent = String(S.cols.length);
      D.navHistCount.textContent = String(S.hist.length);
    }
    function countTypes(P) {
      const m = {};
      P.forEach(p => { m[p.type] = (m[p.type] || 0) + 1; });
      return Object.keys(m).filter(k => m[k]).map(k => ({ t: k, n: m[k] }));
    }

    /* =========================================================
       载入完成后的联动
       ========================================================= */
    function afterLoad() {
      reprofile();
      computeIdx();
      enableNav(true);
      updateSourceCard();
      renderOverview();
      fillColumnSelects();
      autoConfigChart();
      renderTable();
      renderChartTypes();
      renderPalettes();
      syncCfgUI();
      renderChart();
      renderPivotUI();
      renderCleanUI();
      addRecent();
      maybeAutosave();
      if (currentView === 'ingest') go('overview');
      else go(currentView);
    }

    function clearAll(silent) {
      S.cols = []; S.rows = []; S.origin = null; S.profile = null; S.issues = [];
      S.hist = []; S.future = []; S.saved = []; S.sel.clear(); S.lastSheet = null;
      S.view = { search: '', tokens: [], sort: { c: -1, d: 0 }, filters: {}, hidden: [], page: 1, size: S.view.size, density: S.view.density, idx: [], stickyFirst: false };
      S.file = { name: '', size: 0, sheets: [], sheet: '', at: 0, source: '' };
      S.pv.last = null; S.searchCache = null;
      if (S.chart) { try { S.chart.dispose(); } catch (e) { } S.chart = null; }
      S.chartOpt = null;
      D.tblHead.innerHTML = ''; D.tblBody.innerHTML = ''; D.pvWrap.innerHTML = '';
      D.kpiGrid.innerHTML = ''; D.profBody.innerHTML = ''; D.issueList.innerHTML = ''; D.numQuick.innerHTML = '';
      D.savedGrid.innerHTML = ''; D.histList.innerHTML = ''; D.cleanDiff.innerHTML = '';
      D.fileInput.value = '';
      enableNav(false);
      updateSourceCard();
      try { localStorage.removeItem('datalens.data'); } catch (e) { }
      go('ingest');
      renderRecent();
      if (!silent) toast('已清空当前数据集', { kind: 'ok', action: '撤销', onAction: () => toast('清空操作不可撤销，请重新导入', { kind: 'info' }) });
    }

    /* =========================================================
       数据概览渲染
       ========================================================= */
    function kpiCard(ic, label, value, foot, tone) {
      return '<div class="kpi"><div class="row between"><span class="kpi-label">' + esc(label) + '</span>' +
        '<span class="kpi-ic" style="' + (tone ? 'background:' + tone.bg + ';color:' + tone.fg : '') + '"><svg class="ic ic-sm"><use href="#' + (tone ? tone.ic || ic : ic) + '"/></svg></span></div>' +
        '<div class="kpi-value num">' + value + '</div>' + (foot ? '<div class="kpi-foot">' + foot + '</div>' : '') + '</div>';
    }

    function renderOverview() {
      if (!S.profile) return;
      const P = S.profile, rows = S.rows.length, colsN = S.cols.length;
      const cells = rows * colsN;
      const miss = P.reduce((s, p) => s + p.blank, 0);
      const missPct = cells ? miss / cells * 100 : 0;
      const dup = countDupRows();
      const nums = P.filter(p => p.type === 'number');
      const mix = countTypes(P);
      const orig = S.origin ? S.origin.rows.length : rows;
      const delta = rows - orig;

      D.ovSubtitle.textContent = S.file.name + ' · 共 ' + fmtNum(rows) + ' 行 ' + colsN + ' 列，已自动识别 ' + mix.map(m => TYPE_LABEL[m.t] + ' ' + m.n + ' 列').join('、');
      D.kpiGrid.innerHTML = [
        kpiCard('i-rows', '数据行数', fmtNum(rows), delta ? '较原始数据 ' + (delta > 0 ? '+' : '') + fmtNum(delta) + ' 行' : '未做任何清洗修改'),
        kpiCard('i-columns', '数据列数', String(colsN), mix.map(m => '<span class="tag-type ' + TYPE_CLS[m.t] + '">' + TYPE_LABEL[m.t] + ' ' + m.n + '</span>').join(' ')),
        kpiCard('i-database', '单元格总数', fmtCompact(cells), '文件 ' + (S.file.size ? fmtBytes(S.file.size) : '—')),
        kpiCard('i-target', '缺失值', fmtNum(miss), '<div class="meter ' + (missPct > 20 ? 'err' : missPct > 5 ? 'warn' : 'ok') + '" style="margin-top:5px"><i style="width:' + clamp(missPct, 0, 100).toFixed(1) + '%"></i></div>占比 ' + missPct.toFixed(1) + '%'),
        kpiCard('i-copy', '重复行', fmtNum(dup), dup ? '建议一键去重' : '没有完全重复的行', dup ? { bg: 'var(--warn-soft)', fg: 'var(--warn)', ic: 'i-copy' } : { bg: 'var(--ok-soft)', fg: 'var(--ok)', ic: 'i-check' }),
        kpiCard('i-clock', '导入时间', new Date(S.file.at || Date.now()).toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' }), new Date(S.file.at || Date.now()).toLocaleDateString('zh-CN') + ' · ' + esc(S.file.source || '本地'))
      ].join('');

      renderProfile(curProfSort());
      renderIssues();
      renderNumQuick(nums);
    }

    let profSortMode = 'order';
    function curProfSort() { return profSortMode; }
    function renderProfile(mode) {
      const P = (S.profile || []).slice();
      const idx = P.map((p, i) => i);
      if (mode === 'missing') idx.sort((a, b) => P[b].missing - P[a].missing);
      else if (mode === 'unique') idx.sort((a, b) => P[b].unique - P[a].unique);
      else if (mode === 'type') idx.sort((a, b) => String(P[a].type).localeCompare(String(P[b].type)) || a - b);
      D.profCount.textContent = P.length + ' 列 · 点击任意列查看完整统计';
      D.profBody.innerHTML = idx.map(k => {
        const p = P[k];
        const missPct = (p.missing * 100).toFixed(p.missing > 0 ? 1 : 0);
        const dist = p.dist && p.dist.length ? p.dist : [];
        const mx = dist.length ? Math.max.apply(null, dist) : 0;
        const spark = dist.slice(0, 16).map(v => '<i style="height:' + (mx ? Math.max(6, v / mx * 100) : 6) + '%"></i>').join('');
        let stat1 = '—', stat2 = '—', stat3 = '—', stat4 = '—';
        if (p.type === 'number' && p.count) {
          stat1 = '<span class="mono">' + fmtCompact(p.min) + ' ~ ' + fmtCompact(p.max) + '</span>';
          stat2 = '<span class="mono">' + fmtCompact(p.mean) + '</span>';
          stat3 = '<span class="mono">' + fmtCompact(p.median) + '</span>';
          stat4 = p.top.length ? p.top.map(t => esc(t[0])).join('、') : '<span class="dim">合计 ' + fmtCompact(p.sum) + '</span>';
        } else if (p.type === 'date' && p.dmin !== null && p.dmin !== undefined) {
          stat1 = '<span class="mono">' + fmtDate(new Date(p.dmin)) + '</span>';
          stat2 = '<span class="mono">' + fmtDate(new Date(p.dmax)) + '</span>';
          stat3 = '<span class="mono">' + (p.spanDays || 0) + ' 天</span>';
          stat4 = '<span class="dim">日期跨度</span>';
        } else {
          stat1 = '<span class="mono">' + (p.unique ? fmtNum(p.unique) : 0) + '</span>';
          stat2 = '<span class="mono">' + p.avgLen.toFixed(1) + ' 字</span>';
          stat3 = '<span class="dim">' + (p.unique === 1 ? '常量' : p.unique === p.total - p.blank ? '唯一' : '分类') + '</span>';
          stat4 = p.top.length ? p.top.map(t => esc(String(t[0]).slice(0, 10)) + '(' + t[1] + ')').join(' ') : '—';
        }
        return '<div class="prof-row" data-col="' + p.i + '" role="button" tabindex="0" style="cursor:pointer">' +
          '<span class="row gap-2 truncate"><svg class="ic ic-sm" style="color:var(--text-3)"><use href="#' + TYPE_IC[p.type] + '"/></svg>' +
          '<b class="truncate">' + esc(p.name) + '</b><span class="tag-type ' + TYPE_CLS[p.type] + '">' + TYPE_LABEL[p.type] + '</span></span>' +
          '<span class="row gap-2"><span class="meter ' + (p.missing > .3 ? 'err' : p.missing > .08 ? 'warn' : 'ok') + '" style="width:34px"><i style="width:' + clamp(p.missing * 100, p.missing > 0 ? 4 : 0, 100) + '%"></i></span><span class="mono xs">' + missPct + '%</span></span>' +
          '<span class="row gap-2"><span class="spark" style="width:62px">' + spark + '</span><span class="xs dim">唯一 ' + fmtNum(p.unique) + '</span></span>' +
          '<span class="opt-col tiny truncate">' + stat1 + '</span>' +
          '<span class="opt-col tiny truncate">' + stat2 + '</span>' +
          '<span class="opt-col tiny truncate">' + stat3 + '</span>' +
          '<span class="opt-col tiny truncate dim">' + stat4 + '</span>' +
          '</div>';
      }).join('') || '<div class="empty"><p>暂无列信息</p></div>';
    }

    function renderIssues() {
      const I = S.issues || [];
      const tone = { err: ['var(--err-soft)', 'var(--err)'], warn: ['var(--warn-soft)', 'var(--warn)'], info: ['var(--info-soft)', 'var(--info)'], ok: ['var(--ok-soft)', 'var(--ok)'] };
      D.issueCount.textContent = I.filter(i => i.lv !== 'ok').length + ' 项';
      D.issueList.innerHTML = I.map((it, k) => {
        const t = tone[it.lv] || tone.info;
        return '<div class="issue"><span class="i-ic" style="background:' + t[0] + ';color:' + t[1] + '"><svg class="ic ic-sm"><use href="#' + it.ic + '"/></svg></span>' +
          '<div class="grow" style="min-width:0"><div class="tiny b">' + esc(it.t) + '</div><div class="tiny dim" style="margin-top:2px">' + esc(it.d) + '</div></div>' +
          (it.act ? '<button class="btn btn-xs btn-soft" data-issue="' + k + '">' + esc(it.act.l) + '</button>' : '') + '</div>';
      }).join('');
      D.issueList.querySelectorAll('[data-issue]').forEach(b => {
        b.onclick = () => { const it = I[+b.dataset.issue]; if (it && it.act) it.act.f(); };
      });
    }

    function renderNumQuick(nums) {
      if (!nums || !nums.length) {
        D.numQuick.innerHTML = '<div class="empty" style="padding:22px"><div class="empty-ic" style="width:48px;height:48px"><svg class="ic"><use href="#i-hash"/></svg></div><p class="tiny">没有识别到数值列，无法计算统计量。<br>可在「数据清洗」中把文本列转换为数值。</p></div>';
        return;
      }
      D.numQuick.innerHTML = nums.map(p => {
        const range = (p.max - p.min) || 1;
        const posMean = clamp((p.mean - p.min) / range * 100, 0, 100);
        return '<div style="padding:9px 0;border-bottom:1px dashed var(--border)">' +
          '<div class="row between gap-2"><span class="tiny b truncate">' + esc(p.name) + '</span>' +
          '<span class="tiny mono dim">σ ' + fmtCompact(p.std) + '</span></div>' +
          '<div class="row between tiny dim mono" style="margin:5px 0 3px"><span>' + fmtCompact(p.min) + '</span><span>均值 ' + fmtCompact(p.mean) + '</span><span>' + fmtCompact(p.max) + '</span></div>' +
          '<div class="meter" style="position:relative"><i style="width:' + posMean.toFixed(1) + '%"></i></div>' +
          '<div class="row between tiny dim" style="margin-top:4px"><span>中位 ' + fmtCompact(p.median) + '</span><span>合计 ' + fmtCompact(p.sum) + '</span></div>' +
          '</div>';
      }).join('');
    }

    /** 列详情抽屉 */
    function openColDrawer(i) {
      const p = S.profile[i]; if (!p) return;
      const dist = p.dist || [];
      const mx = dist.length ? Math.max.apply(null, dist) : 0;
      const lines = [];
      const add = (k, v) => lines.push('<div class="stat-line"><span class="muted">' + k + '</span><b>' + v + '</b></div>');
      add('数据类型', '<span class="tag-type ' + TYPE_CLS[p.type] + '">' + TYPE_LABEL[p.type] + '</span>');
      add('非空值', fmtNum(p.total - p.blank) + ' / ' + fmtNum(p.total));
      add('缺失值', fmtNum(p.blank) + '（' + (p.missing * 100).toFixed(1) + '%）');
      add('唯一值', fmtNum(p.unique));
      if (p.type === 'number' && p.count) {
        add('最小值', fmtNum(p.min)); add('25% 分位', fmtNum(p.q1)); add('中位数', fmtNum(p.median));
        add('75% 分位', fmtNum(p.q3)); add('最大值', fmtNum(p.max));
        add('平均值', fmtNum(p.mean, 2)); add('标准差', fmtNum(p.std, 2)); add('合计', fmtNum(p.sum, 2));
        add('离群值(3×IQR)', fmtNum(p.outliers || 0));
      } else if (p.type === 'date' && p.dmin !== null && p.dmin !== undefined) {
        add('最早', fmtDate(new Date(p.dmin))); add('最晚', fmtDate(new Date(p.dmax))); add('跨度', (p.spanDays || 0) + ' 天');
      } else {
        add('平均长度', p.avgLen.toFixed(1) + ' 字符');
      }
      if (p.unparse) add('无法解析', fmtNum(p.unparse) + ' 个值');
      add('示例值', esc(String(fmtVal(p.sample, p.type)).slice(0, 40)) || '—');

      const topHtml = p.top && p.top.length
        ? '<div class="card" style="margin-top:12px"><div class="card-h"><h3 class="tiny">高频取值 Top ' + p.top.length + '</h3></div><div class="card-b tight">' +
        p.top.map(t => '<div class="row between gap-2 tiny" style="padding:4px 0"><span class="truncate">' + esc(String(t[0]).slice(0, 24)) + '</span><span class="mono dim">' + fmtNum(t[1]) + '</span></div>').join('') + '</div></div>'
        : '';
      const distHtml = dist.length ? '<div class="card" style="margin-top:12px"><div class="card-h"><h3 class="tiny">' + (p.type === 'number' ? '数值分布（16 组）' : '高频值占比') + '</h3></div>' +
        '<div class="card-b"><div style="display:flex;align-items:flex-end;gap:3px;height:96px">' +
        dist.map(v => '<div style="flex:1;height:' + (mx ? Math.max(3, v / mx * 100) : 3) + '%;background:var(--grad);border-radius:3px 3px 0 0" data-tip="' + fmtNum(v) + '"></div>').join('') +
        '</div><div class="row between tiny dim" style="margin-top:6px"><span>' + (p.type === 'number' ? fmtCompact(p.min) : '高频') + '</span><span>' + (p.type === 'number' ? fmtCompact(p.max) : '低频') + '</span></div></div></div>' : '';

      openDrawer(p.name, '<div class="row gap-2 wrap" style="margin-bottom:12px">' +
        '<span class="badge badge-brand">' + TYPE_LABEL[p.type] + '</span>' +
        '<span class="badge">第 ' + (i + 1) + ' 列</span>' +
        (p.missing > 0 ? '<span class="badge badge-warn">缺失 ' + (p.missing * 100).toFixed(1) + '%</span>' : '<span class="badge badge-ok">无缺失</span>') +
        '</div>' + lines.join('') + distHtml + topHtml +
        '<div class="row gap-2 wrap" style="margin-top:16px">' +
        '<button class="btn btn-sm" data-a="sort"><svg class="ic ic-sm"><use href="#i-sort"/></svg>按此列排序</button>' +
        '<button class="btn btn-sm" data-a="chart"><svg class="ic ic-sm"><use href="#i-chart"/></svg>用此列作图</button>' +
        '<button class="btn btn-sm" data-a="filter"><svg class="ic ic-sm"><use href="#i-filter"/></svg>筛选此列</button>' +
        '<button class="btn btn-sm btn-danger" data-a="drop"><svg class="ic ic-sm"><use href="#i-trash"/></svg>删除此列</button>' +
        '</div>', {
        icon: TYPE_IC[p.type],
        onMount(root) {
          root.querySelectorAll('[data-a]').forEach(b => b.onclick = () => {
            const a = b.dataset.a;
            if (a === 'sort') { S.view.sort = { c: i, d: p.type === 'text' ? 1 : -1 }; computeIdx(); renderTable(); go('table'); toast('已按「' + p.name + '」排序', { kind: 'ok' }); }
            if (a === 'chart') { useColForChart(i); }
            if (a === 'filter') { go('table'); setTimeout(() => openFilterPop(i, D.tblHead.querySelector('th[data-c="' + i + '"]')), 120); }
            if (a === 'drop') dropCol(i);
          });
        }
      });
    }

    function useColForChart(i) {
      const p = S.profile[i];
      go('charts');
      if (p.type === 'number') {
        S.cfg.type = S.cfg.type === 'hist' ? 'hist' : 'hist';
        S.cfg.y = [i];
        if (S.cfg.x >= 0 && S.profile[S.cfg.x] && S.profile[S.cfg.x].type === 'number' && S.cfg.x !== i) S.cfg.type = 'scatter';
      } else {
        S.cfg.x = i;
        const firstNum = S.profile.findIndex(q => q.type === 'number');
        if (firstNum >= 0) S.cfg.y = [firstNum];
        S.cfg.type = 'bar';
      }
      syncCfgUI(); renderChart();
      toast('已把「' + p.name + '」设为图表字段', { kind: 'ok' });
    }

/* =========================================================
       数据表：索引计算 / 筛选 / 排序 / 渲染
       ========================================================= */
    S.dataVer = 0;
    function bumpVer() { S.dataVer++; S.searchCache = null; }

    function rowSearchStr(r) {
      let s = '';
      for (let i = 0; i < r.length; i++) {
        const v = r[i];
        s += (v === null || v === undefined) ? '' : (v instanceof Date ? fmtDate(v, true) : String(v));
        s += '\u0001';
      }
      return s.toLowerCase();
    }
    function ensureSearchCache() {
      if (!S.searchCache || S.searchCache.ver !== S.dataVer) {
        const arr = new Array(S.rows.length);
        for (let i = 0; i < S.rows.length; i++) arr[i] = rowSearchStr(S.rows[i]);
        S.searchCache = { ver: S.dataVer, arr };
      }
      return S.searchCache.arr;
    }

    function valKey(v, type) {
      if (isBlank(v)) return '__blank__';
      if (v instanceof Date) return fmtDate(v);
      if (type === 'number') { const n = toNum(v); return isNaN(n) ? String(v) : String(n); }
      return String(v);
    }

    function matchFilter(v, f, ci) {
      const type = S.cols[ci].type;
      if (f.kind === 'vals') return f.set.has(valKey(v, type));
      const op = f.op;
      if (op === 'blank') return isBlank(v);
      if (op === 'notblank') return !isBlank(v);
      if (isBlank(v)) return false;
      if (f.type === 'number') {
        const n = toNum(v); if (isNaN(n)) return false;
        const a = toNum(f.a), b = toNum(f.b);
        switch (op) {
          case 'gt': return n > a; case 'gte': return n >= a;
          case 'lt': return n < a; case 'lte': return n <= a;
          case 'eq': return n === a; case 'ne': return n !== a;
          case 'between': return n >= Math.min(a, b) && n <= Math.max(a, b);
        }
        return true;
      }
      if (f.type === 'date') {
        const d = toDate(v); if (!d) return false;
        const t = d.getTime();
        const a = toDate(f.a), b = toDate(f.b);
        const day = x => x ? new Date(x.getFullYear(), x.getMonth(), x.getDate()).getTime() : NaN;
        switch (op) {
          case 'after': return a ? t >= day(a) + 864e5 : true;
          case 'before': return a ? t < day(a) : true;
          case 'on': return a ? (t >= day(a) && t < day(a) + 864e5) : true;
          case 'between': return (a && b) ? (t >= day(a) && t < day(b) + 864e5) : true;
        }
        return true;
      }
      const s = String(v).toLowerCase(), q = String(f.v === undefined || f.v === null ? '' : f.v).toLowerCase();
      switch (op) {
        case 'contains': return s.indexOf(q) >= 0;
        case 'notcontains': return s.indexOf(q) < 0;
        case 'eq': return s === q;
        case 'ne': return s !== q;
        case 'starts': return s.indexOf(q) === 0;
        case 'ends': return q && s.slice(-q.length) === q;
        case 'regex': try { return new RegExp(q, 'i').test(s); } catch (e) { return false; }
      }
      return true;
    }

    function computeIdx() {
      const rows = S.rows, tok = S.view.tokens, fl = S.view.filters;
      const fkeys = Object.keys(fl).map(Number).filter(k => fl[k] && (fl[k].kind === 'vals' ? fl[k].set.size : true));
      const cache = tok.length ? ensureSearchCache() : null;
      const idx = [];
      for (let i = 0; i < rows.length; i++) {
        if (tok.length) {
          const s = cache[i]; let ok = true;
          for (let t = 0; t < tok.length; t++) if (s.indexOf(tok[t]) < 0) { ok = false; break; }
          if (!ok) continue;
        }
        if (fkeys.length) {
          let ok = true;
          const r = rows[i];
          for (let k = 0; k < fkeys.length; k++) { const ci = fkeys[k]; if (!matchFilter(r[ci], fl[ci], ci)) { ok = false; break; } }
          if (!ok) continue;
        }
        idx.push(i);
      }
      const srt = S.view.sort;
      if (srt.c >= 0 && srt.d !== 0 && S.cols[srt.c]) {
        const c = srt.c, type = S.cols[c].type, dir = srt.d;
        const keys = new Array(rows.length);
        for (let i = 0; i < rows.length; i++) keys[i] = sortKey(rows[i][c], type);
        idx.sort((a, b) => {
          const ka = keys[a], kb = keys[b];
          if (ka === null && kb === null) return a - b;
          if (ka === null) return 1;
          if (kb === null) return -1;
          let r;
          if (typeof ka === 'number' && typeof kb === 'number') r = ka - kb;
          else r = String(ka).localeCompare(String(kb), 'zh-CN', { numeric: true, sensitivity: 'base' });
          return r * dir;
        });
      }
      S.view.idx = idx;
      const pages = totalPages();
      if (S.view.page > pages) S.view.page = Math.max(1, pages);
      return idx;
    }
    function totalPages() {
      const size = S.view.size;
      if (!size) return 1;
      return Math.max(1, Math.ceil(S.view.idx.length / size));
    }

    function visibleCols() {
      const out = [];
      for (let i = 0; i < S.cols.length; i++) if (S.view.hidden.indexOf(i) < 0) out.push(i);
      return out;
    }

    function renderTable() {
      if (!S.rows.length) {
        D.tblWrap.classList.add('hidden');
        D.tblEmpty.innerHTML = emptyBox('i-table', '还没有数据', '请先在「数据接入」中导入 CSV / Excel / JSON，或载入示例数据。', '<button class="btn btn-primary btn-sm" onclick="go(\'ingest\')"><svg class="ic ic-sm"><use href="#i-upload"/></svg>去导入</button>');
        return;
      }
      D.tblEmpty.innerHTML = '';
      D.tblWrap.classList.remove('hidden');
      const cols = visibleCols();
      const idx = S.view.idx;
      const size = S.view.size || idx.length || 1;
      const pages = totalPages();
      const start = (S.view.page - 1) * size;
      const pageIdx = idx.slice(start, size ? start + size : undefined);
      const tok = S.view.tokens;
      const srt = S.view.sort;
      const fl = S.view.filters;

      let h = '<tr><th class="col-chk"><div class="th-in center"><input class="cbx" type="checkbox" id="chkAll" aria-label="全选本页"></div></th>' +
        '<th class="col-idx"><div class="th-in center">#</div></th>';
      cols.forEach(ci => {
        const c = S.cols[ci], p = S.profile ? S.profile[ci] : null;
        const cls = 'th-sort' + (srt.c === ci ? (srt.d === 1 ? ' asc' : srt.d === -1 ? ' desc' : '') : '') + (fl[ci] ? ' has-filter' : '');
        h += '<th data-c="' + ci + '" class="' + cls + '" aria-sort="' + (srt.c === ci && srt.d === 1 ? 'ascending' : srt.c === ci && srt.d === -1 ? 'descending' : 'none') + '">' +
          '<div class="th-in"><span class="truncate">' + esc(c.name) + '</span>' +
          '<span class="tag-type ' + TYPE_CLS[c.type] + '">' + TYPE_LABEL[c.type] + '</span>' +
          '<svg class="ic ic-sm sarr"><use href="#i-arrow-up"/></svg>' +
          '<span class="th-tools">' +
          '<button class="fbtn" data-f="' + ci + '" aria-label="筛选此列" data-tip="筛选"><svg class="ic ic-sm"><use href="#i-filter"/></svg></button>' +
          '<button data-h="' + ci + '" aria-label="隐藏此列" data-tip="隐藏列"><svg class="ic ic-sm"><use href="#i-eye-off"/></svg></button>' +
          '</span></div></th>';
      });
      h += '</tr>';
      D.tblHead.innerHTML = h;

      const frag = [];
      if (!pageIdx.length) {
        frag.push('<tr><td colspan="' + (cols.length + 2) + '"><div class="empty" style="padding:34px">' +
          '<div class="empty-ic"><svg class="ic ic-lg"><use href="#i-search"/></svg></div>' +
          '<h4>没有匹配的行</h4><p>试着放宽搜索关键词，或清除部分筛选条件。</p>' +
          '<button class="btn btn-sm" style="margin-top:12px" onclick="clearAllFilters()">清除全部筛选</button></div></td></tr>');
      } else {
        for (let k = 0; k < pageIdx.length; k++) {
          const ri = pageIdx[k], r = S.rows[ri];
          const sel = S.sel.has(ri);
          let tr = '<tr data-r="' + ri + '"' + (sel ? ' class="sel"' : '') + '>' +
            '<td class="col-chk"><div class="th-in center" style="padding:0"><input class="cbx" type="checkbox" data-sel="' + ri + '"' + (sel ? ' checked' : '') + ' aria-label="选择第 ' + (ri + 1) + ' 行"></div></td>' +
            '<td class="col-idx" data-row="' + ri + '" title="点击查看整行详情" style="cursor:pointer"><div class="td-in center mono" style="color:var(--text-3);font-size:11.5px">' + (start + k + 1) + '</div></td>';
          for (let j = 0; j < cols.length; j++) {
            const ci = cols[j], c = S.cols[ci], v = r[ci];
            const blank = isBlank(v);
            let cls = 'td-in', txt;
            if (blank) { cls += ' cell-null'; txt = '—'; }
            else if (c.type === 'number') { const n = toNum(v); cls += ' cell-num'; txt = isNaN(n) ? hlText(String(v), tok) : fmtNum(n); }
            else if (c.type === 'date') { const d = toDate(v); txt = d ? hlText(fmtVal(d, 'date'), tok) : hlText(String(v), tok); }
            else if (c.type === 'bool') { txt = /^(true|1|是|yes|y|对)$/i.test(String(v).trim()) ? '是' : '否'; }
            else { cls += ' cell-txt'; txt = hlText(String(v), tok); }
            tr += '<td data-c="' + ci + '" data-r="' + ri + '"><div class="' + cls + '" title="' + (blank ? '空值' : esc(String(v instanceof Date ? fmtDate(v, true) : v)).slice(0, 120)) + '">' + txt + '</div></td>';
          }
          frag.push(tr + '</tr>');
        }
      }
      D.tblBody.innerHTML = frag.join('');
      D.tbl.className = 'tbl d-' + S.view.density;

      D.tblRange.textContent = '显示 ' + (idx.length ? start + 1 : 0) + '-' + (start + pageIdx.length) + ' / 共 ' + fmtNum(idx.length) + ' 行' +
        (idx.length !== S.rows.length ? '（已过滤 ' + fmtNum(S.rows.length - idx.length) + ' 行）' : '');
      D.tblCount.textContent = idx.length === S.rows.length ? fmtNum(idx.length) + ' 行' : fmtNum(idx.length) + ' / ' + fmtNum(S.rows.length) + ' 行';
      D.pageJump.max = String(pages); D.pageJump.value = String(S.view.page);
      renderPager(pages);
      renderChips();
      const chk = document.getElementById('chkAll');
      if (chk) {
        const all = pageIdx.length > 0 && pageIdx.every(ri => S.sel.has(ri));
        chk.checked = all; chk.indeterminate = !all && pageIdx.some(ri => S.sel.has(ri));
      }
      updateSelUI();
      D.tblSubtitle.innerHTML = '点击表头排序 · 双击单元格编辑 · 右键打开单元格菜单 · 点击 <span class="mono">#</span> 列查看整行详情';
    }

    function emptyBox(ic, title, desc, btn) {
      return '<div class="empty"><div class="empty-ic"><svg class="ic ic-lg"><use href="#' + ic + '"/></svg></div><h4>' + esc(title) + '</h4><p>' + esc(desc) + '</p>' + (btn ? '<div style="margin-top:14px">' + btn + '</div>' : '') + '</div>';
    }

    function renderPager(pages) {
      const cur = S.view.page;
      const btn = (label, page, dis, on, tip) => '<button class="pg' + (on ? ' on' : '') + '" data-p="' + page + '"' + (dis ? ' disabled' : '') + (tip ? ' data-tip="' + tip + '"' : '') + '>' + label + '</button>';
      let h = btn('<svg class="ic ic-sm"><use href="#i-chev-l"/></svg><svg class="ic ic-sm" style="margin-left:-6px"><use href="#i-chev-l"/></svg>', 1, cur === 1, false, '第一页');
      h += btn('<svg class="ic ic-sm"><use href="#i-chev-l"/></svg>', cur - 1, cur === 1, false, '上一页');
      const win = [];
      const add = p => { if (p >= 1 && p <= pages && win.indexOf(p) < 0) win.push(p); };
      add(1); for (let p = cur - 1; p <= cur + 1; p++) add(p); add(pages);
      win.sort((a, b) => a - b);
      let prev = 0;
      win.forEach(p => { if (p - prev > 1) h += '<span class="dim tiny" style="padding:0 2px">…</span>'; h += btn(String(p), p, false, p === cur); prev = p; });
      h += btn('<svg class="ic ic-sm"><use href="#i-chev-r"/></svg>', cur + 1, cur === pages, false, '下一页');
      h += btn('<svg class="ic ic-sm"><use href="#i-chev-r"/></svg><svg class="ic ic-sm" style="margin-left:-6px"><use href="#i-chev-r"/></svg>', pages, cur === pages, false, '最后一页');
      h += '<span class="tiny dim" style="margin-left:6px">共 ' + pages + ' 页</span>';
      D.pager.innerHTML = h;
    }

    function updateSelUI() {
      const n = S.sel.size;
      D.btnDelSel.classList.toggle('hidden', n === 0);
      D.selCount.textContent = String(n);
    }

    /* ---------- 筛选条件描述 & chips ---------- */
    const OP_TEXT = { contains: '包含', notcontains: '不包含', eq: '等于', ne: '不等于', starts: '开头是', ends: '结尾是', regex: '正则匹配', blank: '为空', notblank: '不为空' };
    const OP_NUM = { gt: '>', gte: '≥', lt: '<', lte: '≤', eq: '=', ne: '≠', between: '介于', blank: '为空', notblank: '不为空' };
    const OP_DATE = { after: '晚于', before: '早于', on: '等于', between: '介于', blank: '为空', notblank: '不为空' };
    function describeFilter(ci, f) {
      const name = S.cols[ci] ? S.cols[ci].name : ('列' + (ci + 1));
      if (f.kind === 'vals') {
        const n = f.set.size, hasBlank = f.set.has('__blank__');
        const list = Array.from(f.set).filter(x => x !== '__blank__').slice(0, 2);
        return name + ' ∈ ' + (list.map(x => x.length > 10 ? x.slice(0, 10) + '…' : x).join('/') || (hasBlank ? '空值' : '')) + (n > 2 ? ' 等' + n + '项' : (hasBlank && list.length ? '+空值' : ''));
      }
      const map = f.type === 'number' ? OP_NUM : f.type === 'date' ? OP_DATE : OP_TEXT;
      const op = map[f.op] || f.op;
      if (f.op === 'blank' || f.op === 'notblank') return name + ' ' + op;
      if (f.op === 'between') return name + ' ' + op + ' ' + f.a + ' ~ ' + f.b;
      return name + ' ' + op + ' ' + (f.v !== undefined ? f.v : f.a);
    }
    function renderChips() {
      const fl = S.view.filters;
      const keys = Object.keys(fl).map(Number).filter(k => fl[k]);
      if (!keys.length && !S.view.search) { D.filterChips.innerHTML = ''; return; }
      let h = '';
      if (S.view.search) h += '<span class="chip"><svg class="ic ic-sm"><use href="#i-search"/></svg>搜索：' + esc(S.view.search) + '<button data-cs aria-label="清除搜索"><svg class="ic ic-sm" style="width:13px;height:13px"><use href="#i-x"/></svg></button></span>';
      keys.forEach(ci => {
        h += '<span class="chip"><svg class="ic ic-sm"><use href="#i-filter"/></svg>' + esc(describeFilter(ci, fl[ci])) +
          '<button data-cf="' + ci + '" aria-label="清除此筛选"><svg class="ic ic-sm" style="width:13px;height:13px"><use href="#i-x"/></svg></button></span>';
      });
      if (keys.length > 1 || (keys.length && S.view.search)) h += '<button class="btn btn-xs btn-ghost" data-clearall>清除全部</button>';
      D.filterChips.innerHTML = h;
      const cs = D.filterChips.querySelector('[data-cs]');
      if (cs) cs.onclick = () => { D.tblSearch.value = ''; applySearch(''); };
      D.filterChips.querySelectorAll('[data-cf]').forEach(b => b.onclick = () => { delete S.view.filters[+b.dataset.cf]; computeIdx(); S.view.page = 1; renderTable(); });
      const ca = D.filterChips.querySelector('[data-clearall]');
      if (ca) ca.onclick = clearAllFilters;
    }
    function clearAllFilters() {
      S.view.filters = {}; D.tblSearch.value = ''; applySearch('');
      computeIdx(); S.view.page = 1; renderTable();
      toast('已清除全部搜索与筛选条件', { kind: 'ok' });
    }
    function applySearch(v) {
      S.view.search = v || '';
      S.view.tokens = S.view.search.trim().toLowerCase().split(/\s+/).filter(Boolean);
      D.btnSearchClear.classList.toggle('hidden', !S.view.search);
      S.view.page = 1;
      computeIdx(); renderTable();
    }

/* =========================================================
       数据表：交互（排序 / 筛选 / 编辑 / 选择 / 导出）
       ========================================================= */
    const scheduleReprofile = debounce(() => {
      if (!S.rows.length) return;
      reprofile(); renderOverview(); updateSourceCard(); buildReportIfActive();
    }, 900);

    function bindTable() {
      D.tblHead.addEventListener('click', e => {
        const fbtn = e.target.closest('[data-f]');
        if (fbtn) { e.stopPropagation(); openFilterPop(+fbtn.dataset.f, fbtn.closest('th')); return; }
        const hbtn = e.target.closest('[data-h]');
        if (hbtn) { e.stopPropagation(); toggleCol(+hbtn.dataset.h, false); return; }
        const th = e.target.closest('th[data-c]');
        if (!th) return;
        const ci = +th.dataset.c, s = S.view.sort;
        if (s.c !== ci) S.view.sort = { c: ci, d: S.cols[ci].type === 'text' ? 1 : -1 };
        else if (s.d === -1) S.view.sort = { c: ci, d: 1 };
        else if (s.d === 1) S.view.sort = { c: ci, d: 0 };
        else S.view.sort = { c: -1, d: 0 };
        computeIdx(); renderTable();
      });
      D.tblHead.addEventListener('change', e => {
        if (e.target.id !== 'chkAll') return;
        const size = S.view.size || S.view.idx.length;
        const start = (S.view.page - 1) * size;
        const pageIdx = S.view.idx.slice(start, size ? start + size : undefined);
        pageIdx.forEach(ri => { if (e.target.checked) S.sel.add(ri); else S.sel.delete(ri); });
        renderTable();
      });

      D.tblBody.addEventListener('click', e => {
        const cb = e.target.closest('[data-sel]');
        if (cb) {
          const ri = +cb.dataset.sel;
          if (cb.checked) S.sel.add(ri); else S.sel.delete(ri);
          const tr = cb.closest('tr'); tr && tr.classList.toggle('sel', cb.checked);
          const chk = document.getElementById('chkAll');
          if (chk) { const size = S.view.size || S.view.idx.length; const st = (S.view.page - 1) * size;
            const pi = S.view.idx.slice(st, size ? st + size : undefined);
            chk.checked = pi.length > 0 && pi.every(x => S.sel.has(x));
            chk.indeterminate = !chk.checked && pi.some(x => S.sel.has(x)); }
          updateSelUI(); return;
        }
        const idxCell = e.target.closest('td.col-idx[data-row]');
        if (idxCell) { openRowDrawer(+idxCell.dataset.row); return; }
        const td = e.target.closest('td[data-c]');
        if (td) { const tr = td.closest('tr'); if (tr) { tr.classList.add('sel'); S.sel.add(+tr.dataset.r); updateSelUI(); } }
      });

      D.tblBody.addEventListener('dblclick', e => {
        const td = e.target.closest('td[data-c]');
        if (td && !td.querySelector('input')) startEdit(td);
      });

      D.tblBody.addEventListener('contextmenu', e => {
        const td = e.target.closest('td[data-c]');
        if (!td) return;
        e.preventDefault();
        openCellMenu(e.clientX, e.clientY, +td.dataset.r, +td.dataset.c);
      });

      D.pager.addEventListener('click', e => {
        const b = e.target.closest('[data-p]');
        if (!b || b.disabled) return;
        S.view.page = clamp(+b.dataset.p, 1, totalPages());
        renderTable();
        D.tblWrap.scrollTop = 0;
      });
      D.pageJump.addEventListener('change', () => {
        S.view.page = clamp(parseInt(D.pageJump.value, 10) || 1, 1, totalPages());
        renderTable();
      });
      D.pageSize.addEventListener('change', () => {
        S.view.size = parseInt(D.pageSize.value, 10) || 0;
        S.set.pageSize = S.view.size; saveSettings(); S.view.page = 1; renderTable();
      });
      D.tblDensity.addEventListener('click', e => {
        const b = e.target.closest('[data-d]'); if (!b) return;
        S.view.density = b.dataset.d; S.set.density = b.dataset.d; saveSettings();
        $$('#tblDensity button').forEach(x => x.classList.toggle('on', x === b));
        renderTable();
      });

      D.tblSearch.addEventListener('input', debounce(() => applySearch(D.tblSearch.value), 220));
      D.tblSearch.addEventListener('keydown', e => { if (e.key === 'Escape') { D.tblSearch.value = ''; applySearch(''); } });
      D.btnSearchClear.onclick = () => { D.tblSearch.value = ''; applySearch(''); D.tblSearch.focus(); };
      D.btnAddFilter.onclick = () => openColPickerPop(D.btnAddFilter, ci => openFilterPop(ci, D.tblHead.querySelector('th[data-c="' + ci + '"]') || D.btnAddFilter));
      D.btnCols.onclick = () => openColsMenu();
      D.btnExportTbl.onclick = () => openExportMenu(D.btnExportTbl);
      D.btnDelSel.onclick = async () => {
        const n = S.sel.size;
        if (!n) return;
        const ok = await confirmDlg({ title: '删除所选行', message: '将删除 ' + n + ' 行数据，此操作可通过「撤销」恢复。', okText: '删除 ' + n + ' 行', danger: true });
        if (!ok) return;
        snapshot('删除 ' + n + ' 行所选数据');
        const keep = [];
        for (let i = 0; i < S.rows.length; i++) if (!S.sel.has(i)) keep.push(S.rows[i]);
        S.rows = keep; S.sel.clear(); bumpVer(); S.dirty = true;
        reprofile(); computeIdx(); renderTable(); renderOverview(); updateSourceCard(); renderHist();
        toast('已删除 ' + n + ' 行', { kind: 'ok', action: '撤销', onAction: undo });
      };
    }

    function startEdit(td) {
      const ci = +td.dataset.c, ri = +td.dataset.r;
      const v = S.rows[ri][ci];
      const raw = v === null || v === undefined ? '' : (v instanceof Date ? fmtDate(v, true) : String(v));
      td.innerHTML = '<input class="cell-edit" value="' + esc(raw) + '" aria-label="编辑 ' + esc(S.cols[ci].name) + '">';
      const inp = td.querySelector('input');
      inp.focus(); inp.select();
      let done = false;
      const commit = ok => {
        if (done) return; done = true;
        if (ok && inp.value !== raw) { applyEdit(ri, ci, inp.value); toast('已更新第 ' + (ri + 1) + ' 行「' + S.cols[ci].name + '」', { kind: 'ok', ms: 2000 }); }
        renderTable();
      };
      inp.onkeydown = ev => {
        if (ev.key === 'Enter') { ev.preventDefault(); commit(true); }
        else if (ev.key === 'Escape') { ev.preventDefault(); commit(false); }
        else if (ev.key === 'Tab') { ev.preventDefault(); commit(true); }
      };
      inp.onblur = () => commit(true);
    }

    function applyEdit(ri, ci, raw) {
      const type = S.cols[ci].type, t = String(raw).trim();
      let v = raw;
      if (t === '') v = '';
      else if (type === 'number') { const n = toNum(t); v = isNaN(n) ? raw : n; }
      else if (type === 'date') { const d = toDate(t); v = d || raw; }
      else if (type === 'bool') v = /^(是|true|1|yes|y|对)$/i.test(t) ? '是' : '否';
      S.rows[ri][ci] = v;
      bumpVer(); S.dirty = true;
      scheduleReprofile();
      maybeAutosave();
    }

    function openCellMenu(x, y, ri, ci) {
      const v = S.rows[ri][ci];
      const p = document.createElement('div');
      p.className = 'pop';
      const item = (ic, label, k, extra) => '<button class="pop-item" data-k="' + k + '"><svg class="ic ic-sm"><use href="#' + ic + '"/></svg>' + esc(label) + (extra || '') + '</button>';
      p.innerHTML = '<div class="pop-head">' + esc(S.cols[ci].name) + ' · 第 ' + (ri + 1) + ' 行</div>' +
        item('i-copy', '复制单元格', 'copy') +
        item('i-pencil', '编辑单元格', 'edit') +
        item('i-rows', '复制整行（可粘贴到 Excel）', 'copyrow') +
        item('i-eye', '查看整行详情', 'row') +
        '<div class="pop-sep"></div>' +
        item('i-filter', isBlank(v) ? '筛选：该列为空' : '筛选：等于「' + String(fmtVal(v, S.cols[ci].type)).slice(0, 14) + '」', 'eq') +
        item('i-x', isBlank(v) ? '排除：该列为空' : '排除此值', 'ne') +
        item('i-sort', '按此列排序（升序）', 'sortasc') +
        item('i-sort', '按此列排序（降序）', 'sortdesc') +
        '<div class="pop-sep"></div>' +
        item('i-chart', '用此列生成图表', 'chart') +
        item('i-eye-off', '隐藏此列', 'hide');
      document.body.appendChild(p);
      const pw = p.offsetWidth, ph = p.offsetHeight;
      p.style.left = clamp(x, 8, window.innerWidth - pw - 8) + 'px';
      p.style.top = clamp(y, 8, window.innerHeight - ph - 8) + 'px';
      curPop = p;
      const away = e => { if (!p.contains(e.target)) close(); };
      const close = () => { p.remove(); document.removeEventListener('mousedown', away, true); document.removeEventListener('keydown', key, true); if (curPop === p) curPop = null; };
      const key = e => { if (e.key === 'Escape') close(); };
      setTimeout(() => { document.addEventListener('mousedown', away, true); document.addEventListener('keydown', key, true); }, 0);
      p.querySelectorAll('[data-k]').forEach(b => b.onclick = async () => {
        const k = b.dataset.k; close();
        const txt = v instanceof Date ? fmtDate(v, true) : String(v === null || v === undefined ? '' : v);
        if (k === 'copy') { await copyText(txt); toast('已复制：' + (txt.slice(0, 30) || '（空值）'), { kind: 'ok', ms: 1800 }); }
        if (k === 'copyrow') { await copyText(S.rows[ri].map(c => c instanceof Date ? fmtDate(c, true) : (c === null || c === undefined ? '' : String(c))).join('\t')); toast('整行已复制为制表符分隔文本', { kind: 'ok', ms: 1800 }); }
        if (k === 'edit') { const td = D.tblBody.querySelector('td[data-c="' + ci + '"][data-r="' + ri + '"]'); if (td) startEdit(td); else toast('该行不在当前页，请先定位到该行', { kind: 'warn' }); }
        if (k === 'row') openRowDrawer(ri);
        if (k === 'eq' || k === 'ne') {
          const type = S.cols[ci].type;
          if (type === 'number' && !isBlank(v)) S.view.filters[ci] = { kind: 'cond', type: 'number', op: k === 'eq' ? 'eq' : 'ne', a: toNum(v) };
          else if (type === 'date' && !isBlank(v)) S.view.filters[ci] = { kind: 'cond', type: 'date', op: k === 'eq' ? 'on' : 'before', a: fmtDate(toDate(v)) };
          else S.view.filters[ci] = isBlank(v)
            ? { kind: 'cond', type: 'text', op: k === 'eq' ? 'blank' : 'notblank' }
            : { kind: 'vals', set: new Set(k === 'eq' ? [valKey(v, type)] : allKeysExcept(ci, valKey(v, type))) };
          computeIdx(); S.view.page = 1; renderTable();
          toast('已按「' + describeFilter(ci, S.view.filters[ci]) + '」筛选', { kind: 'ok' });
        }
        if (k === 'sortasc') { S.view.sort = { c: ci, d: 1 }; computeIdx(); renderTable(); }
        if (k === 'sortdesc') { S.view.sort = { c: ci, d: -1 }; computeIdx(); renderTable(); }
        if (k === 'chart') useColForChart(ci);
        if (k === 'hide') toggleCol(ci, false);
      });
    }
    function allKeysExcept(ci, key) {
      const set = new Set(), type = S.cols[ci].type;
      for (const r of S.rows) { const k = valKey(r[ci], type); if (k !== key) set.add(k); }
      return set;
    }

    function openRowDrawer(ri) {
      if (ri < 0 || ri >= S.rows.length) return;
      const r = S.rows[ri];
      const body = '<dl style="margin:0">' + S.cols.map((c, ci) => {
        const v = r[ci], blank = isBlank(v);
        return '<div class="kv"><dt>' + esc(c.name) + ' <span class="tag-type ' + TYPE_CLS[c.type] + '">' + TYPE_LABEL[c.type] + '</span></dt>' +
          '<dd class="' + (blank ? 'cell-null' : '') + '">' + (blank ? '空值' : esc(fmtVal(v, c.type))) + '</dd></div>';
      }).join('') + '</dl>' +
        '<div class="row gap-2 wrap" style="margin-top:16px">' +
        '<button class="btn btn-sm" data-a="copy"><svg class="ic ic-sm"><use href="#i-copy"/></svg>复制整行</button>' +
        '<button class="btn btn-sm" data-a="edit"><svg class="ic ic-sm"><use href="#i-pencil"/></svg>编辑该行</button>' +
        '<button class="btn btn-sm btn-danger" data-a="del"><svg class="ic ic-sm"><use href="#i-trash"/></svg>删除该行</button>' +
        '</div>';
      const nav = '<button class="btn btn-ghost btn-icon btn-sm" data-prev aria-label="上一行"><svg class="ic ic-sm"><use href="#i-chev-l"/></svg></button>' +
        '<button class="btn btn-ghost btn-icon btn-sm" data-next aria-label="下一行"><svg class="ic ic-sm"><use href="#i-chev-r"/></svg></button>';
      const d = openDrawer('第 ' + (ri + 1) + ' 行详情', body, { icon: 'i-rows', actions: nav });
      d.root.querySelector('[data-prev]').onclick = () => { d.close(); openRowDrawer(ri - 1); };
      d.root.querySelector('[data-next]').onclick = () => { d.close(); openRowDrawer(ri + 1); };
      d.root.querySelector('[data-a="copy"]').onclick = async () => {
        await copyText(S.cols.map(c => c.name).join('\t') + '\n' + r.map(v => v instanceof Date ? fmtDate(v, true) : (v === null || v === undefined ? '' : String(v))).join('\t'));
        toast('已复制（含表头）', { kind: 'ok', ms: 1800 });
      };
      d.root.querySelector('[data-a="edit"]').onclick = () => {
        d.close(); go('table');
        const pos = S.view.idx.indexOf(ri);
        if (pos < 0) { toast('该行已被当前筛选条件过滤掉', { kind: 'warn' }); return; }
        const size = S.view.size || S.view.idx.length || 1;
        S.view.page = Math.floor(pos / size) + 1;
        renderTable();
        setTimeout(() => {
          const td = D.tblBody.querySelector('td[data-c="' + visibleCols()[0] + '"][data-r="' + ri + '"]');
          if (td) { td.scrollIntoView({ block: 'center' }); startEdit(td); }
        }, 80);
      };
      d.root.querySelector('[data-a="del"]').onclick = async () => {
        d.close();
        const ok = await confirmDlg({ title: '删除该行', message: '将删除第 ' + (ri + 1) + ' 行，可通过撤销恢复。', okText: '删除', danger: true });
        if (!ok) return;
        snapshot('删除第 ' + (ri + 1) + ' 行');
        S.rows.splice(ri, 1); S.sel.delete(ri); bumpVer(); S.dirty = true;
        reprofile(); computeIdx(); renderTable(); renderOverview(); updateSourceCard(); renderHist();
        toast('已删除该行', { kind: 'ok', action: '撤销', onAction: undo });
      };
    }

    function toggleCol(ci, show) {
      const h = S.view.hidden;
      const at = h.indexOf(ci);
      if (show === undefined) show = at >= 0;
      if (show && at >= 0) h.splice(at, 1);
      if (!show && at < 0) {
        if (visibleCols().length <= 1) { toast('至少要保留一列可见', { kind: 'warn' }); return; }
        h.push(ci);
      }
      renderTable();
    }

    function openColsMenu() {
      const html = '<div class="pop-head">显示 / 隐藏列</div>' +
        '<div class="row gap-2" style="padding:2px 8px 8px">' +
        '<button class="btn btn-xs" data-all="1">全选</button><button class="btn btn-xs" data-all="0">全不选</button>' +
        '<button class="btn btn-xs" data-all="num">仅数值列</button></div>' +
        '<div style="max-height:320px;overflow:auto">' +
        S.cols.map((c, i) => '<label class="valitem"><input class="cbx" type="checkbox" data-ci="' + i + '"' + (S.view.hidden.indexOf(i) < 0 ? ' checked' : '') + '>' +
          '<svg class="ic ic-sm" style="color:var(--text-3)"><use href="#' + TYPE_IC[c.type] + '"/></svg>' +
          '<span class="grow">' + esc(c.name) + '</span></label>').join('') + '</div>';
      openPop(D.btnCols, html, {
        onMount(p) {
          p.querySelectorAll('[data-ci]').forEach(cb => cb.onchange = () => { toggleCol(+cb.dataset.ci, cb.checked); refreshColsMenuState(p); });
          p.querySelectorAll('[data-all]').forEach(b => b.onclick = () => {
            const m = b.dataset.all;
            S.view.hidden = m === '1' ? [] : S.cols.map((c, i) => i).filter(i => !(m === 'num' ? S.cols[i].type === 'number' : false));
            if (m === '0') S.view.hidden = S.cols.map((c, i) => i).slice(0, -1);
            renderTable();
            p.querySelectorAll('[data-ci]').forEach(cb => cb.checked = S.view.hidden.indexOf(+cb.dataset.ci) < 0);
          });
        }
      });
    }
    function refreshColsMenuState(p) { p.querySelectorAll('[data-ci]').forEach(cb => cb.checked = S.view.hidden.indexOf(+cb.dataset.ci) < 0); }

    function openColPickerPop(anchor, cb) {
      const html = '<div class="pop-head">选择列</div><div style="max-height:320px;overflow:auto">' +
        S.cols.map((c, i) => '<button class="pop-item" data-ci="' + i + '"><svg class="ic ic-sm" style="color:var(--text-3)"><use href="#' + TYPE_IC[c.type] + '"/></svg>' + esc(c.name) +
          '<span class="grow"></span><span class="tag-type ' + TYPE_CLS[c.type] + '">' + TYPE_LABEL[c.type] + '</span></button>').join('') + '</div>';
      openPop(anchor, html, { onMount(p, close) { p.querySelectorAll('[data-ci]').forEach(b => b.onclick = () => { close(); cb(+b.dataset.ci); }); } });
    }

    function openExportMenu(anchor) {
      const nAll = S.rows.length, nF = S.view.idx.length, nS = S.sel.size;
      const html = '<div class="pop-head">导出当前数据</div>' +
        '<button class="pop-item" data-e="csv|filtered"><svg class="ic ic-sm"><use href="#i-file"/></svg>CSV · 筛选结果<span class="grow"></span><span class="tiny dim">' + fmtNum(nF) + ' 行</span></button>' +
        '<button class="pop-item" data-e="csv|all"><svg class="ic ic-sm"><use href="#i-file"/></svg>CSV · 全部数据<span class="grow"></span><span class="tiny dim">' + fmtNum(nAll) + ' 行</span></button>' +
        (nS ? '<button class="pop-item" data-e="csv|sel"><svg class="ic ic-sm"><use href="#i-check"/></svg>CSV · 所选行<span class="grow"></span><span class="tiny dim">' + fmtNum(nS) + ' 行</span></button>' : '') +
        '<div class="pop-sep"></div>' +
        '<button class="pop-item" data-e="xlsx|filtered"><svg class="ic ic-sm"><use href="#i-table"/></svg>Excel (.xlsx) · 筛选结果</button>' +
        '<button class="pop-item" data-e="xlsx|all"><svg class="ic ic-sm"><use href="#i-table"/></svg>Excel (.xlsx) · 全部数据</button>' +
        '<div class="pop-sep"></div>' +
        '<button class="pop-item" data-e="json|filtered"><svg class="ic ic-sm"><use href="#i-braces"/></svg>JSON · 对象数组</button>' +
        '<button class="pop-item" data-e="md|filtered"><svg class="ic ic-sm"><use href="#i-report"/></svg>Markdown 表格（前 200 行）</button>';
      openPop(anchor, html, {
        onMount(p, close) {
          p.querySelectorAll('[data-e]').forEach(b => b.onclick = () => { close(); const a = b.dataset.e.split('|'); exportData(a[0], a[1]); });
        }
      });
    }

    function exportRows(scope) {
      let rows;
      if (scope === 'all') rows = S.rows;
      else if (scope === 'sel') rows = S.view.idx.filter(i => S.sel.has(i)).length ? S.view.idx.filter(i => S.sel.has(i)).map(i => S.rows[i]) : Array.from(S.sel).sort((a, b) => a - b).map(i => S.rows[i]);
      else rows = S.view.idx.map(i => S.rows[i]);
      return rows;
    }
    function exportData(fmt, scope) {
      if (!S.rows.length) { toast('没有可导出的数据', { kind: 'warn' }); return; }
      const rows = exportRows(scope);
      const cols = S.cols;
      if (!rows.length) { toast('当前范围内没有数据行', { kind: 'warn' }); return; }
      const base = (S.file.name || 'datalens').replace(/\.[^.]+$/, '') + (S.dirty ? '_已清洗' : '');
      const plain = rows.map(r => r.map(v => v instanceof Date ? fmtDate(v, true) : (isBlank(v) ? '' : v)));
      try {
        if (fmt === 'csv') downloadText('\uFEFF' + toCSV(cols.map(c => c.name), plain), stampName(base, 'csv'), 'text/csv');
        else if (fmt === 'json') {
          const objs = rows.map(r => { const o = {}; cols.forEach((c, i) => { const v = r[i]; o[c.name] = v instanceof Date ? fmtDate(v, true) : (isBlank(v) ? null : (c.type === 'number' ? (isNaN(toNum(v)) ? v : toNum(v)) : v)); }); return o; });
          downloadText(JSON.stringify(objs, null, 2), stampName(base, 'json'), 'application/json');
        } else if (fmt === 'md') {
          const cs = visibleCols().slice(0, 12), rs = rows.slice(0, 200);
          const md = '| ' + cs.map(i => esc(cols[i].name)).join(' | ') + ' |\n| ' + cs.map(() => '---').join(' | ') + ' |\n' +
            rs.map(r => '| ' + cs.map(i => String(fmtVal(r[i], cols[i].type)).replace(/\|/g, '\\|').slice(0, 40)).join(' | ') + ' |').join('\n');
          downloadText(md, stampName(base, 'md'), 'text/markdown');
        } else if (fmt === 'xlsx') {
          if (!window.XLSX) { toast('Excel 导出需要 SheetJS 库，请联网后重试', { kind: 'err' }); return; }
          const aoa = [cols.map(c => c.name)].concat(rows.map(r => r.map(v => v instanceof Date ? fmtDate(v, true) : (isBlank(v) ? '' : v))));
          const ws = XLSX.utils.aoa_to_sheet(aoa);
          ws['!cols'] = cols.map((c, i) => ({ wch: clamp(Math.max(c.name.length * 2, 10), 8, 40) }));
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, '数据');
          XLSX.writeFile(wb, stampName(base, 'xlsx'));
        }
        toast('已导出 ' + fmtNum(rows.length) + ' 行 · ' + fmt.toUpperCase(), { kind: 'ok', title: '导出成功' });
      } catch (e) {
        console.error(e);
        toast(String(e.message || e), { kind: 'err', title: '导出失败' });
      }
    }

/* =========================================================
       列筛选弹层
       ========================================================= */
    function distinctValues(ci, limit) {
      const m = new Map(), type = S.cols[ci].type;
      for (let i = 0; i < S.rows.length; i++) {
        const k = valKey(S.rows[i][ci], type);
        m.set(k, (m.get(k) || 0) + 1);
        if (m.size > (limit || 300) * 4) break;
      }
      const arr = Array.from(m.entries()).sort((a, b) => b[1] - a[1]);
      const blank = arr.filter(e => e[0] === '__blank__');
      const rest = arr.filter(e => e[0] !== '__blank__');
      return blank.concat(rest).slice(0, limit || 300);
    }

    function openFilterPop(ci, anchor) {
      const type = S.cols[ci].type;
      const cur = S.view.filters[ci] || null;
      let kind = cur ? cur.kind : (type === 'text' || type === 'bool' ? 'vals' : 'cond');
      const vals = distinctValues(ci, 300);
      const sel = cur && cur.kind === 'vals' ? new Set(cur.set) : new Set();
      const opSel = (map, curOp) => Object.keys(map).map(k => '<option value="' + k + '"' + (curOp === k ? ' selected' : '') + '>' + map[k] + '</option>').join('');

      const html =
        '<div class="pop-head" style="display:flex;align-items:center;gap:6px">' +
        '<svg class="ic ic-sm"><use href="#' + TYPE_IC[type] + '"/></svg>' + esc(S.cols[ci].name) +
        '<span class="tag-type ' + TYPE_CLS[type] + '" style="margin-left:auto">' + TYPE_LABEL[type] + '</span></div>' +
        '<div class="seg" style="width:calc(100% - 12px);margin:2px 6px 8px" data-tabs>' +
        '<button data-t="vals" class="grow center' + (kind === 'vals' ? ' on' : '') + '">按值选择</button>' +
        '<button data-t="cond" class="grow center' + (kind === 'cond' ? ' on' : '') + '">按条件</button></div>' +

        '<div data-p="vals"' + (kind === 'vals' ? '' : ' hidden') + '>' +
        '<div style="padding:0 8px 6px" class="row gap-2">' +
        '<input class="input input-sm" data-vsearch placeholder="搜索取值…" style="flex:1">' +
        '<button class="btn btn-xs" data-vall>全选</button><button class="btn btn-xs" data-vnone>清空</button></div>' +
        '<div style="max-height:236px;overflow:auto;padding:0 4px" data-vlist>' +
        vals.map(e => '<label class="valitem" data-lbl="' + esc(e[0]) + '"><input class="cbx" type="checkbox" data-v="' + esc(e[0]) + '"' + (sel.has(e[0]) ? ' checked' : '') + '>' +
          '<span class="grow truncate">' + (e[0] === '__blank__' ? '<i class="dim">（空值）</i>' : esc(String(e[0]).slice(0, 34))) + '</span>' +
          '<span class="cnt">' + fmtNum(e[1]) + '</span></label>').join('') +
        '</div><div class="hint" style="padding:6px 9px 0">共 ' + vals.length + ' 个取值' + (vals.length >= 300 ? '（仅显示前 300）' : '') + ' · 不勾选任何项表示不限制</div>' +
        '</div>' +

        '<div data-p="cond"' + (kind === 'cond' ? '' : ' hidden') + ' style="padding:2px 9px 6px" class="col gap-2">' +
        '<select class="select input-sm" data-op>' + opSel(type === 'number' ? OP_NUM : type === 'date' ? OP_DATE : OP_TEXT, cur ? cur.op : (type === 'number' ? 'gte' : type === 'date' ? 'after' : 'contains')) + '</select>' +
        '<input class="input input-sm" data-a placeholder="' + (type === 'date' ? 'YYYY-MM-DD' : '值') + '" value="' + esc(cur && cur.kind === 'cond' ? (cur.v !== undefined ? cur.v : cur.a) : '') + '"' + (type === 'date' ? ' type="date"' : '') + '>' +
        '<input class="input input-sm" data-b placeholder="至（区间用）" value="' + esc(cur && cur.kind === 'cond' && cur.op === 'between' ? cur.b : '') + '"' + (type === 'date' ? ' type="date"' : '') + '>' +
        '<div class="hint" id="fCondHint"></div>' +
        '</div>' +

        '<div class="pop-sep"></div>' +
        '<div class="row gap-2" style="padding:2px 6px 4px">' +
        '<button class="btn btn-xs btn-danger" data-clear>清除</button><div class="grow"></div>' +
        '<button class="btn btn-xs" data-cancel>取消</button>' +
        '<button class="btn btn-xs btn-primary" data-apply>应用筛选</button></div>';

      openPop(anchor || D.btnAddFilter, html, {
        align: 'left',
        onMount(p, close) {
          const tabs = p.querySelector('[data-tabs]');
          const panV = p.querySelector('[data-p="vals"]'), panC = p.querySelector('[data-p="cond"]');
          const op = p.querySelector('[data-op]'), ia = p.querySelector('[data-a]'), ib = p.querySelector('[data-b]');
          const hint = p.querySelector('#fCondHint');
          const syncHint = () => {
            const o = op.value;
            ib.style.display = o === 'between' ? '' : 'none';
            ia.style.display = (o === 'blank' || o === 'notblank') ? 'none' : '';
            const n = S.profile && S.profile[ci] ? S.profile[ci] : null;
            hint.textContent = n && type === 'number' && n.count ? '该列范围：' + fmtNum(n.min) + ' ~ ' + fmtNum(n.max) + '，中位数 ' + fmtNum(n.median) : (n && type === 'text' ? '唯一值 ' + fmtNum(n.unique) + ' 个' : '');
          };
          op.onchange = syncHint; syncHint();
          tabs.onclick = e => {
            const b = e.target.closest('[data-t]'); if (!b) return;
            kind = b.dataset.t;
            $$('.seg button', tabs).forEach(x => x.classList.toggle('on', x === b));
            panV.hidden = kind !== 'vals'; panC.hidden = kind !== 'cond';
          };
          const vs = p.querySelector('[data-vsearch]');
          vs.oninput = () => {
            const q = vs.value.trim().toLowerCase();
            p.querySelectorAll('[data-lbl]').forEach(l => {
              l.style.display = !q || String(l.dataset.lbl).toLowerCase().indexOf(q) >= 0 ? '' : 'none';
            });
          };
          p.querySelector('[data-vall]').onclick = () => p.querySelectorAll('[data-v]').forEach(c => { if (c.closest('[data-lbl]').style.display !== 'none') c.checked = true; });
          p.querySelector('[data-vnone]').onclick = () => p.querySelectorAll('[data-v]').forEach(c => c.checked = false);
          p.querySelector('[data-cancel]').onclick = close;
          p.querySelector('[data-clear]').onclick = () => {
            delete S.view.filters[ci]; computeIdx(); S.view.page = 1; renderTable(); close();
            toast('已清除「' + S.cols[ci].name + '」的筛选', { kind: 'ok', ms: 1800 });
          };
          p.querySelector('[data-apply]').onclick = () => {
            if (kind === 'vals') {
              const set = new Set();
              p.querySelectorAll('[data-v]').forEach(c => { if (c.checked) set.add(c.dataset.v); });
              if (!set.size || set.size >= vals.length) delete S.view.filters[ci];
              else S.view.filters[ci] = { kind: 'vals', set };
            } else {
              const o = op.value;
              if (o === 'blank' || o === 'notblank') S.view.filters[ci] = { kind: 'cond', type, op: o };
              else if (type === 'number') {
                const a = toNum(ia.value), b = toNum(ib.value);
                if (o === 'between' && (isNaN(a) || isNaN(b))) { toast('请输入有效的区间数值', { kind: 'warn' }); return; }
                if (o !== 'between' && isNaN(a)) { toast('请输入有效的数值', { kind: 'warn' }); return; }
                S.view.filters[ci] = { kind: 'cond', type: 'number', op: o, a: isNaN(a) ? 0 : a, b: isNaN(b) ? 0 : b };
              } else if (type === 'date') {
                if (!ia.value) { toast('请选择日期', { kind: 'warn' }); return; }
                S.view.filters[ci] = { kind: 'cond', type: 'date', op: o, a: ia.value, b: ib.value };
              } else {
                if (!ia.value && o !== 'regex') { toast('请输入筛选内容', { kind: 'warn' }); return; }
                S.view.filters[ci] = { kind: 'cond', type: 'text', op: o, v: ia.value };
              }
            }
            computeIdx(); S.view.page = 1; renderTable(); close();
            const f = S.view.filters[ci];
            toast(f ? '筛选后剩 ' + fmtNum(S.view.idx.length) + ' 行' : '已取消该列筛选', { kind: 'ok', title: f ? describeFilter(ci, f) : '筛选已清除' });
          };
          setTimeout(() => { if (kind === 'cond') ia.focus(); else vs.focus(); }, 50);
        }
      });
    }

    /* =========================================================
       列下拉框填充（图表 / 透视 / 清洗共用）
       ========================================================= */
    function optHtml(ci, extra) {
      const c = S.cols[ci];
      return '<option value="' + ci + '"' + (extra || '') + '>' + esc(c.name) + '（' + TYPE_LABEL[c.type] + '）</option>';
    }
    function fillColumnSelects() {
      const n = S.cols.length;
      const all = S.cols.map((c, i) => optHtml(i)).join('');
      const nums = S.cols.map((c, i) => c.type === 'number' ? optHtml(i) : '').join('');
      const cats = S.cols.map((c, i) => (c.type !== 'number') ? optHtml(i) : '').join('');

      D.cfgX.innerHTML = '<option value="-1">（不分组 / 逐行）</option>' + all;
      D.cfgSplit.innerHTML = '<option value="-1">不拆分</option>' + all;

      D.yList.innerHTML = S.cols.map((c, i) => '<label class="yitem"><input class="cbx" type="checkbox" data-y="' + i + '"' +
        (c.type === 'number' ? '' : ' title="非数值列将按计数处理"') + '>' +
        '<svg class="ic ic-sm" style="color:var(--text-3)"><use href="#' + TYPE_IC[c.type] + '"/></svg>' +
        '<span class="grow">' + esc(c.name) + '</span>' +
        '<span class="tag-type ' + TYPE_CLS[c.type] + '">' + TYPE_LABEL[c.type] + '</span></label>').join('');

      D.pvCol.innerHTML = '<option value="-1">（不分列）</option>' + all;
      D.pvVal.innerHTML = (nums || all);
      D.pvRowList.innerHTML = S.cols.map((c, i) => '<label class="yitem"><input class="cbx" type="checkbox" data-pvr="' + i + '">' +
        '<svg class="ic ic-sm" style="color:var(--text-3)"><use href="#' + TYPE_IC[c.type] + '"/></svg><span class="grow">' + esc(c.name) + '</span>' +
        '<span class="tag-type ' + TYPE_CLS[c.type] + '">' + TYPE_LABEL[c.type] + '</span></label>').join('');

      D.fillCol.innerHTML = all;
      D.colOp.innerHTML = all;
      D.splitCol.innerHTML = all;
      D.dedupeCol.innerHTML = all;

      // 计算列快捷插入
      D.calcQuick.innerHTML = '<span class="hint" style="margin-right:4px">快速插入：</span>' +
        S.cols.map((c, i) => '<button class="btn btn-xs" data-ins="' + esc(c.name) + '">[' + esc(c.name.length > 10 ? c.name.slice(0, 10) + '…' : c.name) + ']</button>').join('') +
        ' <button class="btn btn-xs" data-ins="IF">IF(条件, A, B)</button><button class="btn btn-xs" data-ins="ROUND">ROUND(x, n)</button>';
      D.calcQuick.querySelectorAll('[data-ins]').forEach(b => b.onclick = () => {
        const t = b.dataset.ins;
        D.calcExpr.value += (/^(IF|ROUND)$/.test(t) ? t + '()' : '[' + t + ']');
        D.calcExpr.focus();
      });
      D.yHint.textContent = '共 ' + n + ' 列，其中数值列 ' + S.cols.filter(c => c.type === 'number').length + ' 个';
    }

    function renderPivotUI() {
      if (!S.cols.length) return;
      D.pvRowList.querySelectorAll('[data-pvr]').forEach(cb => cb.checked = S.pv.rows.indexOf(+cb.dataset.pvr) >= 0);
      D.pvCol.value = String(S.pv.col); D.pvVal.value = String(S.pv.val >= 0 ? S.pv.val : -1);
      if (S.pv.val < 0) { const firstNum = S.cols.findIndex(c => c.type === 'number'); if (firstNum >= 0) D.pvVal.value = String(firstNum); }
      D.pvAgg.value = S.pv.agg;
      renderPivot();
    }

/* =========================================================
       可视化
       ========================================================= */
    function shortLabel(s, n) {
      s = String(s === null || s === undefined ? '' : s);
      if (!s.trim()) return '（空）';
      return s.length > (n || 18) ? s.slice(0, (n || 18) - 1) + '…' : s;
    }
    function isDark() { return document.documentElement.getAttribute('data-theme') === 'dark'; }
    function chartColors() {
      const dark = isDark();
      return {
        pal: PALETTES[S.cfg.palette % PALETTES.length].c,
        text: dark ? '#c3cde0' : '#4a5568',
        title: dark ? '#eef2fa' : '#0f172a',
        axis: dark ? '#2c3752' : '#e4e8f1',
        split: dark ? 'rgba(148,163,184,.14)' : 'rgba(15,23,42,.07)',
        bg: dark ? '#121826' : '#ffffff'
      };
    }

    function renderChartTypes() {
      D.chartTypeGrid.className = 'grid typegrid';
      D.chartTypeGrid.style.gridTemplateColumns = 'repeat(3,minmax(0,1fr))';
      D.chartTypeGrid.innerHTML = CHART_TYPES.map(t =>
        '<button data-t="' + t.id + '" class="' + (S.cfg.type === t.id ? 'on' : '') + '" data-tip="' + t.n + '">' +
        '<svg class="ic"><use href="#' + t.ic + '"/></svg>' + t.n + '</button>').join('');
      D.chartTypeGrid.querySelectorAll('[data-t]').forEach(b => b.onclick = () => {
        S.cfg.type = b.dataset.t;
        D.chartTypeGrid.querySelectorAll('[data-t]').forEach(x => x.classList.toggle('on', x === b));
        applyTypeRules(); renderChart();
      });
    }
    function renderPalettes() {
      D.paletteList.innerHTML = PALETTES.map((p, i) =>
        '<button class="sw ' + (S.cfg.palette === i ? 'on' : '') + '" data-p="' + i + '" data-tip="' + p.name + '" aria-label="配色 ' + esc(p.name) + '" style="background:linear-gradient(135deg,' + p.c[0] + ' 0%,' + p.c[1] + ' 50%,' + p.c[2] + ' 100%)"></button>').join('');
      D.paletteList.querySelectorAll('[data-p]').forEach(b => b.onclick = () => {
        S.cfg.palette = +b.dataset.p; S.set.palette = S.cfg.palette; saveSettings();
        D.paletteList.querySelectorAll('[data-p]').forEach(x => x.classList.toggle('on', x === b));
        renderChart(); renderSaved();
      });
    }

    /** 根据图表类型显示 / 隐藏相关配置项 */
    function applyTypeRules() {
      const t = S.cfg.type;
      const single = ['pie', 'ring', 'funnel'].indexOf(t) >= 0;
      const noDim = ['hist', 'heat'].indexOf(t) >= 0;
      D.wrapSplit.style.display = single || noDim || t === 'scatter' ? 'none' : '';
      D.wrapAgg.style.display = (t === 'hist' || t === 'heat' || S.cfg.mode === 'raw') ? 'none' : '';
      D.wrapTopN.style.display = (t === 'hist' || t === 'heat') ? 'none' : '';
      D.wrapSortBy.style.display = (t === 'hist' || t === 'heat' || S.cfg.mode === 'raw') ? 'none' : '';
      D.wrapBin.style.display = t === 'hist' ? '' : 'none';
      const hint = {
        bar: '适合比较不同类别的数值大小。', barh: '类别名称较长时更易阅读。',
        line: '适合时间序列与趋势变化。', area: '强调累计量与趋势。',
        pie: '展示各部分占整体的比例（建议 ≤ 8 个类别）。', ring: '环形图，中心可放总计。',
        scatter: '探索两个数值变量的关系，维度列作为 X。', radar: '多维度能力对比，取前若干类别。',
        funnel: '按数值降序展示转化漏斗。', box: '查看数值分布与离群值（可按维度分组）。',
        hist: '查看单个数值列的频率分布。', heat: '数值列之间的皮尔逊相关系数矩阵。'
      };
      D.modeHint.textContent = S.cfg.mode === 'agg' ? '把相同维度的行合并后再统计，适合分类汇总。' : '逐行绘制原始数据点，适合小数据集或时间序列。';
      D.chartSubtitle.textContent = hint[t] || '';
      if (single && S.cfg.y.length > 1) S.cfg.y = S.cfg.y.slice(0, 1);
      if (single) { S.cfg.split = -1; D.cfgSplit.value = '-1'; }
      if (t === 'hist' || t === 'heat') S.cfg.mode = 'agg';
      syncYList();
    }

    function syncCfgUI() {
      D.cfgX.value = String(S.cfg.x);
      D.cfgSplit.value = String(S.cfg.split);
      D.cfgAgg.value = S.cfg.agg;
      D.cfgTopN.value = String(S.cfg.topN);
      D.cfgSortBy.value = S.cfg.sortBy;
      D.cfgBin.value = String(S.cfg.bin); D.cfgBinVal.textContent = String(S.cfg.bin);
      D.cfgSmooth.checked = !!S.cfg.smooth; D.cfgStack.checked = !!S.cfg.stack;
      D.cfgLabel.checked = !!S.cfg.label; D.cfgRotate.checked = !!S.cfg.rotate;
      D.cfgDataZoom.checked = !!S.cfg.zoom;
      D.cfgTitle.value = S.cfg.title || '';
      $$('#chartMode button').forEach(b => b.classList.toggle('on', b.dataset.m === S.cfg.mode));
      D.modeBadge.textContent = S.cfg.mode === 'agg' ? '聚合' : '明细';
      D.chartTypeGrid.querySelectorAll('[data-t]').forEach(x => x.classList.toggle('on', x.dataset.t === S.cfg.type));
      D.paletteList.querySelectorAll('[data-p]').forEach(x => x.classList.toggle('on', +x.dataset.p === S.cfg.palette));
      syncYList(); applyTypeRules();
    }
    function syncYList() {
      D.yList.querySelectorAll('[data-y]').forEach(cb => cb.checked = S.cfg.y.indexOf(+cb.dataset.y) >= 0);
    }

    function autoConfigChart(quiet) {
      const P = S.profile; if (!P || !P.length) return null;
      const cfg = S.cfg;
      const dateCol = P.findIndex(p => p.type === 'date' && p.unique > 1);
      let x = -1, why = '';
      if (dateCol >= 0) { x = dateCol; why = '检测到日期列「' + P[dateCol].name + '」，用折线图看趋势'; }
      else {
        const cands = P.filter(p => (p.type === 'text' || p.type === 'bool') && p.unique > 1 && p.unique <= Math.min(60, Math.max(2, S.rows.length * 0.5)));
        cands.sort((a, b) => b.unique - a.unique);
        if (cands.length) { x = cands[cands.length - 1].i; why = '选择分类粒度适中的「' + P[x].name + '」作为维度'; }
        else { const t = P.findIndex(p => p.type === 'text'); if (t >= 0) { x = t; why = '使用文本列「' + P[t].name + '」作为维度'; } }
      }
      const prefer = /(金额|销售|收入|利润|数量|总额|单价|成本|评分|满意度|amount|sales|revenue|profit|qty|price|score|total)/i;
      const nums = P.filter(p => p.type === 'number' && p.count > 0 && p.max !== p.min);
      const pref = nums.filter(p => prefer.test(p.name));
      cfg.x = x;
      cfg.y = (pref.length ? pref : nums).slice(0, 1).map(p => p.i);
      cfg.mode = 'agg';
      cfg.split = -1;
      if (!cfg.y.length) { cfg.agg = 'countAll'; cfg.type = x >= 0 ? 'bar' : 'hist'; why += '；没有合适的数值列，改为统计行数'; }
      else {
        cfg.agg = 'sum';
        cfg.type = dateCol >= 0 && dateCol === x ? 'line' : (x >= 0 ? 'bar' : 'hist');
      }
      cfg.sortBy = cfg.type === 'line' ? 'xasc' : 'desc';
      cfg.topN = cfg.type === 'line' ? 0 : 20;
      cfg.zoom = cfg.type === 'line' && x >= 0 && P[x] && P[x].unique > 40;
      cfg.title = '';
      syncCfgUI(); renderChart();
      if (!quiet) toast(why || '已根据数据结构自动选择维度与指标', { kind: 'ok', title: '智能推荐' });
      return cfg;
    }

    /** 计算绘图数据 */
    function computeChartData() {
      const cfg = S.cfg, rows = S.rows, idx = S.view.idx;
      const out = { cats: [], series: [], warn: '', meta: '' };
      if (!rows.length) return out;
      const yIdx = cfg.y.slice();
      const xIdx = cfg.x, spIdx = cfg.split;
      const xType = xIdx >= 0 ? S.cols[xIdx].type : null;

      /* --- 直方图 --- */
      if (cfg.type === 'hist') {
        const yi = yIdx.length ? yIdx[0] : S.cols.findIndex(c => c.type === 'number');
        if (yi < 0) { out.warn = '请至少选择一个数值列'; return out; }
        const vals = [];
        for (const ri of idx) { const n = toNum(rows[ri][yi]); if (!isNaN(n)) vals.push(n); }
        if (!vals.length) { out.warn = '「' + S.cols[yi].name + '」没有可解析的数值'; return out; }
        const h = histogram(vals, cfg.bin);
        out.cats = h.counts.map((c, i) => fmtCompact(h.bins[i]) + ' ~ ' + fmtCompact(h.bins[i + 1]));
        out.series = [{ name: '频数', data: h.counts }];
        out.meta = '共 ' + fmtNum(vals.length) + ' 个有效值 · ' + h.counts.length + ' 组 · 组宽 ' + fmtCompact(h.width);
        return out;
      }

      /* --- 相关性热力图 --- */
      if (cfg.type === 'heat') {
        let nums = yIdx.map(i => i).filter(i => S.cols[i] && S.cols[i].type === 'number');
        if (nums.length < 2) nums = S.cols.map((c, i) => c.type === 'number' ? i : -1).filter(i => i >= 0).slice(0, 10);
        if (nums.length < 2) { out.warn = '相关性热力图需要至少 2 个数值列'; return out; }
        nums = nums.slice(0, 12);
        const cols = nums.map(i => { const a = []; for (const ri of idx) { const n = toNum(rows[ri][i]); a.push(isNaN(n) ? null : n); } return a; });
        const data = [];
        for (let a = 0; a < nums.length; a++) for (let b = 0; b < nums.length; b++) {
          const xs = [], ys = [];
          for (let k = 0; k < cols[a].length; k++) if (cols[a][k] !== null && cols[b][k] !== null) { xs.push(cols[a][k]); ys.push(cols[b][k]); }
          data.push([b, a, xs.length > 2 ? +pearson(xs, ys).toFixed(3) : 0]);
        }
        out.cats = nums.map(i => shortLabel(S.cols[i].name, 12));
        out.series = [{ name: '相关系数', data, matrix: true }];
        out.meta = nums.length + ' 个数值列 · ' + fmtNum(idx.length) + ' 行参与计算';
        return out;
      }

      /* --- 散点图 --- */
      if (cfg.type === 'scatter') {
        if (!yIdx.length) { out.warn = '请选择至少一个数值指标'; return out; }
        const LIMIT = 4000;
        const use = idx.length > LIMIT ? idx.slice(0, LIMIT) : idx;
        if (idx.length > LIMIT) out.warn = '数据点较多，仅绘制前 ' + fmtNum(LIMIT) + ' 行';
        const splitOn = spIdx >= 0 && S.cols[spIdx];
        const xIsNum = xIdx >= 0 && (S.cols[xIdx].type === 'number' || S.cols[xIdx].type === 'date');
        let xMap = null;
        if (xIdx >= 0 && !xIsNum) { xMap = new Map(); for (const ri of use) { const k = valKey(rows[ri][xIdx], xType); if (!xMap.has(k)) xMap.set(k, xMap.size); } }
        const xv = ri => {
          if (xIdx < 0) return ri + 1;
          const v = rows[ri][xIdx];
          if (xIsNum) { const n = toNum(v); return isNaN(n) ? null : n; }
          return xMap ? xMap.get(valKey(v, xType)) : null;
        };
        const groups = new Map();
        for (const ri of use) {
          const sk = splitOn ? shortLabel(fmtVal(rows[ri][spIdx], S.cols[spIdx].type), 16) : (S.cols[yIdx[0]].name);
          if (!groups.has(sk)) groups.set(sk, []);
          groups.get(sk).push(ri);
        }
        let names = Array.from(groups.keys());
        if (names.length > 8) { names = names.slice(0, 8); }
        out.series = [];
        names.forEach(n => {
          const list = groups.get(n) || [];
          out.series.push({
            name: n, data: list.map(ri => {
              const y = toNum(rows[ri][yIdx[0]]);
              const x = xv(ri);
              return (x === null || isNaN(y)) ? null : [x, y];
            }).filter(Boolean)
          });
        });
        out.cats = xMap ? Array.from(xMap.keys()).map(k => shortLabel(k === '__blank__' ? '（空）' : k, 14)) : null;
        out.meta = fmtNum(use.length) + ' 行 · X：' + (xIdx >= 0 ? S.cols[xIdx].name : '行号') + ' · Y：' + S.cols[yIdx[0]].name;
        return out;
      }

      /* --- 箱线图 --- */
      if (cfg.type === 'box') {
        const yi = yIdx.length ? yIdx : S.cols.map((c, i) => c.type === 'number' ? i : -1).filter(i => i >= 0).slice(0, 3);
        if (!yi.length) { out.warn = '没有可用的数值列'; return out; }
        if (xIdx >= 0) {
          const g = new Map();
          for (const ri of idx) { const k = valKey(rows[ri][xIdx], xType); if (!g.has(k)) g.set(k, []); g.get(k).push(ri); }
          let keys = Array.from(g.entries()).sort((a, b) => b[1].length - a[1].length);
          const topN = cfg.topN > 0 ? cfg.topN : 20;
          if (keys.length > topN) keys = keys.slice(0, topN);
          if (yi.length === 1) {
            out.cats = keys.map(k => shortLabel(k[0] === '__blank__' ? '（空）' : k[0], 16));
            out.series = [{ name: S.cols[yi[0]].name, data: keys.map(k => boxStats(k[1].map(ri => toNum(rows[ri][yi[0]])).filter(n => !isNaN(n)))), box: true }];
          } else {
            out.cats = [];
            out.series = yi.map(y => ({ name: S.cols[y].name, data: [], box: true }));
            keys.forEach(k => {
              out.cats.push(shortLabel(k[0] === '__blank__' ? '（空）' : k[0], 16));
              yi.forEach((y, s) => out.series[s].data.push(boxStats(k[1].map(ri => toNum(rows[ri][y])).filter(n => !isNaN(n)))));
            });
          }
        } else {
          out.cats = yi.map(y => shortLabel(S.cols[y].name, 16));
          out.series = [{
            name: '分布', box: true, data: yi.map(y => {
              const a = []; for (const ri of idx) { const n = toNum(rows[ri][y]); if (!isNaN(n)) a.push(n); }
              return boxStats(a);
            })
          }];
        }
        out.meta = '箱体 = Q1~Q3，中线 = 中位数，须 = 1.5×IQR 内的极值';
        return out;
      }

      /* --- 明细模式 --- */
      if (cfg.mode === 'raw') {
        const LIMIT = 3000;
        const use = idx.length > LIMIT ? idx.slice(0, LIMIT) : idx;
        if (idx.length > LIMIT) out.warn = '明细模式仅绘制前 ' + fmtNum(LIMIT) + ' 行，建议改用聚合模式';
        if (!yIdx.length) { out.warn = '请选择至少一个指标列'; return out; }
        out.cats = use.map((ri, k) => xIdx >= 0 ? shortLabel(fmtVal(rows[ri][xIdx], xType), 16) : '#' + (ri + 1));
        out.series = yIdx.map(yi => ({
          name: S.cols[yi].name,
          data: use.map(ri => { const n = toNum(rows[ri][yi]); return isNaN(n) ? null : n; })
        }));
        out.meta = fmtNum(use.length) + ' 行 · 未聚合';
        return out;
      }

      /* --- 聚合模式 --- */
      if (!yIdx.length) {
        if (cfg.agg !== 'countAll' && cfg.agg !== 'count' && cfg.agg !== 'distinct') { out.warn = '请选择至少一个指标列'; return out; }
      }
      const groups = new Map();
      for (const ri of idx) {
        const xv = xIdx >= 0 ? valKey(rows[ri][xIdx], xType) : '__all__';
        let g = groups.get(xv);
        if (!g) { g = { label: xIdx >= 0 ? shortLabel(fmtVal(rows[ri][xIdx], xType), 20) : '全部数据', items: new Map() }; groups.set(xv, g); }
        const sk = spIdx >= 0 && S.cols[spIdx] ? shortLabel(fmtVal(rows[ri][spIdx], S.cols[spIdx].type), 16) : '__';
        if (!g.items.has(sk)) g.items.set(sk, []);
        g.items.get(sk).push(ri);
      }
      const splitOn = spIdx >= 0 && S.cols[spIdx];
      const splitKeys = [];
      groups.forEach(g => g.items.forEach((v, k) => { if (splitKeys.indexOf(k) < 0) splitKeys.push(k); }));
      const effSplits = splitOn ? splitKeys.slice(0, 10) : ['__'];
      const ys = yIdx.length ? yIdx : [-1];

      let cats = Array.from(groups.entries()).map(e => ({ key: e[0], label: e[1].label, g: e[1] }));
      const seriesOf = (cat, sk, yi) => aggValues(cat.g.items.get(sk) || [], rows, yi, yi < 0 ? 'countAll' : cfg.agg);

      // 汇总值用于排序 / TopN（每个单元只计算一次）
      const vkey = (sk, yi) => sk + '\u0003' + yi;
      cats.forEach(c => {
        c.vals = Object.create(null);
        c.total = 0;
        effSplits.forEach(sk => ys.forEach(yi => {
          const v = seriesOf(c, sk, yi);
          c.vals[vkey(sk, yi)] = v;
          c.total += isNaN(v) ? 0 : Math.abs(v);
        }));
      });

      const sb = cfg.sortBy;
      if (sb === 'desc') cats.sort((a, b) => b.total - a.total);
      else if (sb === 'asc') cats.sort((a, b) => a.total - b.total);
      else if (sb === 'xdesc') cats.sort((a, b) => String(b.label).localeCompare(String(a.label), 'zh-CN', { numeric: true }));
      // xasc 保持原顺序

      const topN = cfg.topN > 0 ? cfg.topN : cats.length;
      let rest = null;
      if (cats.length > topN) {
        rest = cats.slice(topN);
        cats = cats.slice(0, topN);
      }
      out.cats = cats.map(c => c.label);
      out.series = [];
      effSplits.forEach(sk => ys.forEach(yi => {
        const name = splitOn ? (ys.length > 1 ? sk + ' · ' + S.cols[yi].name : sk) : (yi < 0 ? '行数' : S.cols[yi].name);
        out.series.push({ name, data: cats.map(c => c.vals[vkey(sk, yi)]) });
      }));
      // 饼 / 环 图把剩余合并为「其他」
      if (rest && rest.length && ['pie', 'ring', 'funnel'].indexOf(cfg.type) >= 0) {
        const sum = rest.reduce((s, c) => s + c.total, 0);
        out.cats.push('其他（' + rest.length + ' 项）');
        out.series.forEach(sr => sr.data.push(+sum.toFixed(6)));
      }
      out.series.forEach(sr => sr.data = sr.data.map(v => typeof v === 'number' ? +v.toFixed(6) : v));
      out.meta = (xIdx >= 0 ? '维度：' + S.cols[xIdx].name : '不分组') +
        (splitOn ? ' · 拆分：' + S.cols[spIdx].name : '') +
        ' · 指标：' + (yIdx.length ? yIdx.map(i => S.cols[i].name).join('、') : '行数') +
        ' · 聚合：' + AGG_NAMES[yIdx.length ? cfg.agg : 'countAll'] +
        ' · 基于 ' + fmtNum(idx.length) + ' 行';
      return out;
    }

    function boxStats(arr) {
      if (!arr.length) return [0, 0, 0, 0, 0];
      const a = arr.slice().sort((x, y) => x - y);
      const q1 = quantile(a, .25), q2 = quantile(a, .5), q3 = quantile(a, .75);
      const iqr = q3 - q1;
      const lo = q1 - 1.5 * iqr, hi = q3 + 1.5 * iqr;
      let mn = a[0], mx = a[a.length - 1];
      for (const v of a) { if (v >= lo) { mn = v; break; } }
      for (let i = a.length - 1; i >= 0; i--) { if (a[i] <= hi) { mx = a[i]; break; } }
      return [+mn.toFixed(4), +q1.toFixed(4), +q2.toFixed(4), +q3.toFixed(4), +mx.toFixed(4)];
    }

    function aggValues(list, rows, valIdx, agg) {
      if (agg === 'countAll') return list.length;
      if (!list.length) return 0;
      if (agg === 'distinct') { const s = new Set(); for (const i of list) { const v = rows[i][valIdx]; if (!isBlank(v)) s.add(valKey(v, S.cols[valIdx] ? S.cols[valIdx].type : 'text')); } return s.size; }
      if (agg === 'count') { let c = 0; for (const i of list) if (!isBlank(rows[i][valIdx]) && !isNaN(toNum(rows[i][valIdx]))) c++; return c; }
      const nums = [];
      for (const i of list) { const n = toNum(rows[i][valIdx]); if (!isNaN(n)) nums.push(n); }
      if (!nums.length) return 0;
      switch (agg) {
        case 'sum': return nums.reduce((a, b) => a + b, 0);
        case 'avg': return mean(nums);
        case 'min': return nums.reduce((a, b) => a < b ? a : b);
        case 'max': return nums.reduce((a, b) => a > b ? a : b);
        case 'median': return median(nums);
      }
      return 0;
    }

/* =========================================================
       可视化：构建 ECharts 配置
       ========================================================= */
    function autoTitle(res) {
      const cfg = S.cfg;
      if (cfg.type === 'hist') return (cfg.y[0] >= 0 ? S.cols[cfg.y[0]].name : '数值') + ' 分布直方图';
      if (cfg.type === 'heat') return '数值列相关性矩阵';
      const WORD = { sum: '合计', avg: '平均值', count: '计数', countAll: '行数', min: '最小值', max: '最大值', median: '中位数', distinct: '去重数' };
      const tn = CHART_TYPES.filter(t => t.id === cfg.type)[0];
      const ys = cfg.y.length ? (cfg.y.length > 1 ? '（' + cfg.y.map(i => S.cols[i].name).join(' + ') + '）' : S.cols[cfg.y[0]].name) : '';
      const word = cfg.y.length ? (WORD[cfg.agg] || '统计') : '行数';
      const x = cfg.x >= 0 ? S.cols[cfg.x].name : null;
      return (x ? x + ' · ' : '') + ys + word + (tn ? '（' + tn.n + '）' : '');
    }

    function buildOption(res) {
      const cfg = S.cfg, C = chartColors(), t = cfg.type;
      const opt = {
        color: C.pal, backgroundColor: 'transparent',
        animationDuration: 620, animationEasing: 'cubicOut',
        textStyle: { color: C.text, fontSize: 12 },
        tooltip: {
          backgroundColor: isDark() ? 'rgba(20,26,42,.96)' : 'rgba(255,255,255,.97)',
          borderColor: C.axis, borderWidth: 1, padding: [8, 11],
          textStyle: { color: isDark() ? '#e8edf7' : '#0f172a', fontSize: 12 },
          extraCssText: 'border-radius:10px;box-shadow:0 10px 30px rgba(15,23,42,.16);',
          confine: true
        }
      };
      if (cfg.title) opt.title = { text: cfg.title, left: 'center', top: 6, textStyle: { fontSize: 15, fontWeight: 600, color: C.title } };
      const topPad = cfg.title ? 54 : 26;
      const fmtTip = v => (typeof v === 'number' ? fmtNum(v, Math.abs(v) < 10 ? 3 : 2) : v);

      /* 相关性热力图 */
      if (t === 'heat') {
        const d = res.series[0].data;
        opt.tooltip = Object.assign(opt.tooltip, { formatter: p => esc(res.cats[p.value[1]]) + ' × ' + esc(res.cats[p.value[0]]) + '<br>相关系数 <b>' + p.value[2] + '</b>' });
        opt.grid = { left: 10, right: 16, top: topPad, bottom: 74, containLabel: true };
        opt.xAxis = { type: 'category', data: res.cats, splitArea: { show: true }, axisLabel: { rotate: 40, color: C.text, fontSize: 11 }, axisLine: { lineStyle: { color: C.axis } } };
        opt.yAxis = { type: 'category', data: res.cats, splitArea: { show: true }, axisLabel: { color: C.text, fontSize: 11 }, axisLine: { lineStyle: { color: C.axis } } };
        opt.visualMap = {
          min: -1, max: 1, calculable: true, orient: 'horizontal', left: 'center', bottom: 6, itemWidth: 12, itemHeight: 90,
          textStyle: { color: C.text, fontSize: 11 },
          inRange: { color: ['#2563eb', '#93c5fd', isDark() ? '#1b2334' : '#f8fafc', '#fca5a5', '#dc2626'] }
        };
        opt.series = [{ type: 'heatmap', data: d, label: { show: res.cats.length <= 8, fontSize: 10.5, formatter: p => p.value[2] }, emphasis: { itemStyle: { shadowBlur: 8, shadowColor: 'rgba(0,0,0,.3)' } } }];
        return opt;
      }

      /* 饼 / 环 / 漏斗 / 雷达 */
      if (['pie', 'ring', 'funnel', 'radar'].indexOf(t) >= 0) {
        const s0 = res.series[0] || { name: '', data: [] };
        const data = res.cats.map((c, i) => ({ name: c, value: s0.data[i] || 0 }));
        opt.tooltip = Object.assign(opt.tooltip, { trigger: 'item', formatter: p => esc(p.name) + '<br><b>' + fmtTip(p.value) + '</b>' + (p.percent !== undefined ? '（' + p.percent + '%）' : '') });
        if (t === 'radar') {
          const cats = res.cats.slice(0, 10);
          let mx = 0;
          res.series.forEach(sr => sr.data.slice(0, 10).forEach(v => { if (typeof v === 'number' && Math.abs(v) > mx) mx = Math.abs(v); }));
          if (!mx) mx = 1;
          opt.legend = { bottom: 4, left: 'center', textStyle: { color: C.text, fontSize: 11.5 }, icon: 'roundRect', itemWidth: 12, itemHeight: 8 };
          opt.radar = {
            indicator: cats.map(c => ({ name: c, max: mx * 1.18 })), center: ['50%', '52%'], radius: '64%',
            axisName: { color: C.text, fontSize: 11 }, splitLine: { lineStyle: { color: C.split } },
            splitArea: { areaStyle: { color: ['transparent', isDark() ? 'rgba(148,163,184,.05)' : 'rgba(15,23,42,.025)'] } },
            axisLine: { lineStyle: { color: C.axis } }
          };
          opt.series = [{
            type: 'radar', symbolSize: 5, areaStyle: { opacity: .16 }, lineStyle: { width: 2.2 },
            data: res.series.slice(0, 6).map(s => ({ name: s.name, value: s.data.slice(0, 10) }))
          }];
          return opt;
        }
        if (t === 'funnel') {
          opt.legend = { bottom: 4, left: 'center', textStyle: { color: C.text, fontSize: 11.5 }, icon: 'roundRect', itemWidth: 12, itemHeight: 8 };
          opt.series = [{
            type: 'funnel', left: '8%', right: '8%', top: topPad, bottom: 44, minSize: '16%', gap: 3,
            label: { show: cfg.label, position: 'inside', formatter: p => p.name + '  ' + fmtTip(p.value), color: '#fff', fontSize: 11.5 },
            itemStyle: { borderColor: C.bg, borderWidth: 2, borderRadius: 5 },
            data: data.slice().sort((a, b) => b.value - a.value)
          }];
          return opt;
        }
        opt.legend = { type: 'scroll', orient: 'vertical', right: 8, top: 'middle', textStyle: { color: C.text, fontSize: 11.5 }, icon: 'circle', itemWidth: 9, itemHeight: 9 };
        opt.series = [{
          name: s0.name || '数值', type: 'pie', radius: t === 'ring' ? ['44%', '70%'] : '68%',
          center: ['40%', '52%'], avoidLabelOverlap: true, padAngle: t === 'ring' ? 2 : 0,
          itemStyle: { borderColor: C.bg, borderWidth: 2, borderRadius: cfg.smooth ? 7 : 0 },
          label: { show: true, formatter: p => (p.percent >= 4 ? p.name + '\n' + p.percent + '%' : ''), color: C.text, fontSize: 11, lineHeight: 15 },
          labelLine: { length: 10, length2: 10, lineStyle: { color: C.axis } },
          emphasis: { scaleSize: 8, label: { show: true, fontSize: 13, fontWeight: 700 } },
          data
        }];
        if (t === 'ring') {
          const total = data.reduce((s, d) => s + (d.value || 0), 0);
          opt.graphic = [{
            type: 'text', left: '33%', top: '46%',
            style: { text: fmtCompact(total) + '\n总计', textAlign: 'center', fill: C.title, fontSize: 17, fontWeight: 700, lineHeight: 22 }
          }];
        }
        return opt;
      }

      /* 箱线图 */
      if (t === 'box') {
        opt.tooltip = Object.assign(opt.tooltip, {
          trigger: 'item', formatter: p => {
            if (!p.value || !p.value.length) return esc(p.name || '');
            const v = p.value.length === 5 ? p.value : p.value.slice(1);
            return '<b>' + esc(p.name || p.seriesName) + '</b><br>最小 ' + fmtTip(v[0]) + '<br>Q1 ' + fmtTip(v[1]) + '<br>中位 ' + fmtTip(v[2]) + '<br>Q3 ' + fmtTip(v[3]) + '<br>最大 ' + fmtTip(v[4]);
          }
        });
        opt.legend = res.series.length > 1 ? { bottom: 4, textStyle: { color: C.text, fontSize: 11.5 }, icon: 'roundRect', itemWidth: 12, itemHeight: 8 } : undefined;
        opt.grid = { left: 14, right: 20, top: topPad, bottom: res.series.length > 1 ? 56 : 40, containLabel: true };
        opt.xAxis = { type: 'category', data: res.cats, axisLabel: { rotate: cfg.rotate && res.cats.length > 6 ? 38 : 0, color: C.text, fontSize: 11 }, axisLine: { lineStyle: { color: C.axis } } };
        opt.yAxis = { type: 'value', axisLabel: { color: C.text, fontSize: 11, formatter: v => fmtCompact(v) }, splitLine: { lineStyle: { color: C.split } } };
        opt.series = res.series.map(s => ({
          name: s.name, type: 'boxplot', data: s.data, boxWidth: [8, 34],
          itemStyle: { borderWidth: 1.6, borderRadius: cfg.smooth ? 4 : 0 }
        }));
        return opt;
      }

      /* 散点图 */
      if (t === 'scatter') {
        opt.tooltip = Object.assign(opt.tooltip, {
          trigger: 'item', formatter: p => {
            const xl = cfg.x >= 0 ? S.cols[cfg.x].name : '行号';
            return '<b>' + esc(p.seriesName) + '</b><br>' + esc(xl) + '：' + fmtTip(Array.isArray(res.cats) && res.cats[p.value[0]] !== undefined && cfg.x >= 0 && S.cols[cfg.x].type !== 'number' && S.cols[cfg.x].type !== 'date' ? res.cats[p.value[0]] : p.value[0]) +
              '<br>' + esc(S.cols[cfg.y[0]].name) + '：' + fmtTip(p.value[1]);
          }
        });
        opt.legend = res.series.length > 1 ? { type: 'scroll', bottom: 4, textStyle: { color: C.text, fontSize: 11.5 }, icon: 'circle', itemWidth: 9, itemHeight: 9 } : undefined;
        opt.grid = { left: 14, right: 22, top: topPad, bottom: res.series.length > 1 ? 56 : 34, containLabel: true };
        const xIsCat = Array.isArray(res.cats);
        opt.xAxis = xIsCat
          ? { type: 'category', data: res.cats, axisLabel: { rotate: res.cats.length > 8 ? 38 : 0, color: C.text, fontSize: 11 }, axisLine: { lineStyle: { color: C.axis } }, splitLine: { show: false } }
          : { type: 'value', name: cfg.x >= 0 ? S.cols[cfg.x].name : '行号', nameTextStyle: { color: C.text, fontSize: 11 }, axisLabel: { color: C.text, fontSize: 11, formatter: v => fmtCompact(v) }, splitLine: { lineStyle: { color: C.split } } };
        opt.yAxis = { type: 'value', name: cfg.y[0] >= 0 ? S.cols[cfg.y[0]].name : '', nameTextStyle: { color: C.text, fontSize: 11 }, axisLabel: { color: C.text, fontSize: 11, formatter: v => fmtCompact(v) }, splitLine: { lineStyle: { color: C.split } } };
        opt.series = res.series.map(s => ({
          name: s.name, type: 'scatter', data: s.data, symbolSize: cfg.seriesSize || 9,
          itemStyle: { opacity: .74, borderColor: C.bg, borderWidth: .6 },
          emphasis: { focus: 'series', itemStyle: { opacity: 1 } }
        }));
        if (cfg.zoom) opt.dataZoom = [{ type: 'inside' }, { type: 'slider', height: 16, bottom: 6, borderColor: C.axis, textStyle: { color: C.text, fontSize: 10 } }];
        return opt;
      }

      /* 柱 / 条 / 折线 / 面积 / 直方图 */
      const horiz = t === 'barh';
      const catAxis = {
        type: 'category', data: res.cats,
        axisLabel: { rotate: !horiz && cfg.rotate && res.cats.length > 6 ? 40 : 0, color: C.text, fontSize: 11, interval: 0, hideOverlap: true, formatter: v => String(v).length > 14 ? String(v).slice(0, 13) + '…' : v },
        axisLine: { lineStyle: { color: C.axis } }, axisTick: { show: false },
        splitLine: { show: false }
      };
      const valAxis = {
        type: 'value', axisLabel: { color: C.text, fontSize: 11, formatter: v => fmtCompact(v) },
        splitLine: { lineStyle: { color: C.split, type: 'dashed' } }, axisLine: { show: false }, axisTick: { show: false }
      };
      opt.tooltip = Object.assign(opt.tooltip, {
        trigger: 'axis', axisPointer: { type: t === 'line' || t === 'area' ? 'line' : 'shadow', lineStyle: { color: C.axis }, shadowStyle: { color: isDark() ? 'rgba(148,163,184,.09)' : 'rgba(15,23,42,.05)' } },
        formatter: ps => {
          const arr = Array.isArray(ps) ? ps : [ps];
          let s = '<b>' + esc(arr[0].axisValueLabel || arr[0].name) + '</b>';
          arr.forEach(p => { s += '<br>' + (p.marker || '') + ' ' + esc(p.seriesName) + '：<b>' + fmtTip(p.value) + '</b>'; });
          if (arr.length > 1) { const sum = arr.reduce((a, p) => a + (typeof p.value === 'number' ? p.value : 0), 0); if (cfg.stack) s += '<br><span style="opacity:.65">合计：' + fmtTip(sum) + '</span>'; }
          return s;
        }
      });
      opt.legend = res.series.length > 1 ? { type: 'scroll', bottom: 2, left: 'center', textStyle: { color: C.text, fontSize: 11.5 }, icon: 'roundRect', itemWidth: 12, itemHeight: 8, pageIconColor: C.text, pageTextStyle: { color: C.text } } : undefined;
      opt.grid = { left: 12, right: 22, top: topPad, bottom: (res.series.length > 1 ? 42 : 18) + (horiz ? 10 : (cfg.rotate && res.cats.length > 6 ? 26 : 8)) + (cfg.zoom ? 30 : 0), containLabel: true };
      opt.xAxis = horiz ? valAxis : catAxis;
      opt.yAxis = horiz ? Object.assign({}, catAxis, { inverse: cfg.sortBy === 'desc' }) : valAxis;

      const isLine = t === 'line' || t === 'area';
      opt.series = res.series.map((s, si) => {
        const base = { name: s.name, data: s.data, emphasis: { focus: 'series' } };
        if (isLine) {
          base.type = 'line'; base.smooth = !!cfg.smooth; base.symbol = 'circle';
          base.symbolSize = res.cats.length > 60 ? 0 : 6; base.showSymbol = res.cats.length <= 120;
          base.lineStyle = { width: 2.4 };
          base.connectNulls = true;
          if (t === 'area') {
            const c1 = C.pal[si % C.pal.length];
            base.areaStyle = {
              opacity: .85,
              color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1, colorStops: [{ offset: 0, color: c1 + '59' }, { offset: 1, color: c1 + '05' }] }
            };
          }
          if (cfg.stack) base.stack = '合计';
        } else {
          base.type = 'bar'; base.barMaxWidth = 44; base.barMinWidth = 2;
          base.itemStyle = { borderRadius: cfg.smooth ? (horiz ? [0, 6, 6, 0] : [6, 6, 0, 0]) : 0 };
          if (cfg.stack) base.stack = '合计';
          if (res.series.length > 1 && !cfg.stack) base.barGap = '18%';
        }
        if (cfg.label) base.label = { show: true, position: horiz ? 'right' : (isLine ? 'top' : 'top'), fontSize: 10.5, color: C.text, formatter: p => fmtCompact(p.value) };
        return base;
      });
      if (cfg.zoom) opt.dataZoom = [{ type: 'inside' }, { type: 'slider', height: 16, bottom: 4, borderColor: C.axis, fillerColor: isDark() ? 'rgba(124,140,255,.18)' : 'rgba(79,70,229,.12)', textStyle: { color: C.text, fontSize: 10 } }];
      return opt;
    }

    function showChartMsg(msg, ic) {
      D.chartBox.classList.add('hidden');
      D.chartEmpty.innerHTML = emptyBox(ic || 'i-chart', '暂无图表', msg, S.rows.length ? '<button class="btn btn-sm btn-primary" onclick="autoConfigChart()"><svg class="ic ic-sm"><use href="#i-wand"/></svg>智能推荐配置</button>' : '<button class="btn btn-sm btn-primary" onclick="go(\'ingest\')">去导入数据</button>');
      D.chartFoot.innerHTML = esc(msg);
      D.chartTitle.textContent = '图表预览';
    }

    function renderChart() {
      if (!window.echarts) { showChartMsg('图表库（ECharts）未能加载，请检查网络后刷新页面重试。', 'i-alert'); return; }
      if (!S.rows.length) { showChartMsg('请先导入数据，再开始可视化分析。'); return; }
      let res;
      try { res = computeChartData(); }
      catch (e) { console.error(e); showChartMsg('计算图表数据时出错：' + (e.message || e), 'i-alert'); return; }
      if (res.warn || !res.series.length || !res.series.some(s => s.data && s.data.length)) {
        if (S.chart) { try { S.chart.clear(); } catch (e) { } }
        showChartMsg(res.warn || '当前配置没有可绘制的数据，请更换维度或指标。');
        return;
      }
      D.chartEmpty.innerHTML = '';
      D.chartBox.classList.remove('hidden');
      try {
        if (!S.chart || S.chart.isDisposed()) {
          S.chart = echarts.init(D.chartBox, null, { renderer: 'canvas' });
          if (!S.__ro && window.ResizeObserver) {
            S.__ro = new ResizeObserver(throttle(() => { if (S.chart && !S.chart.isDisposed()) S.chart.resize(); }, 160));
            S.__ro.observe(D.chartBox);
          }
        }
        const o = buildOption(res);
        S.chartOpt = o;
        S.chart.setOption(o, true);
        S.chart.resize();
      } catch (e) {
        console.error(e); showChartMsg('渲染图表失败：' + (e.message || e), 'i-alert'); return;
      }
      D.chartTitle.textContent = S.cfg.title || autoTitle(res);
      D.chartFoot.innerHTML = '<span>' + esc(res.meta || '') + '</span>' +
        (res.warn ? '<span class="badge badge-warn" style="margin-left:8px">' + esc(res.warn) + '</span>' : '') +
        '<span class="grow"></span><span class="dim">' + new Date().toLocaleTimeString('zh-CN') + ' 更新</span>';
    }

    const scheduleRenderChart = debounce(() => renderChart(), 260);

    function bindChartUI() {
      D.chartMode.onclick = e => {
        const b = e.target.closest('[data-m]'); if (!b) return;
        S.cfg.mode = b.dataset.m;
        $$('#chartMode button').forEach(x => x.classList.toggle('on', x === b));
        D.modeBadge.textContent = S.cfg.mode === 'agg' ? '聚合' : '明细';
        applyTypeRules(); renderChart();
      };
      D.cfgX.onchange = () => { S.cfg.x = +D.cfgX.value; renderChart(); };
      D.cfgSplit.onchange = () => { S.cfg.split = +D.cfgSplit.value; renderChart(); };
      D.cfgAgg.onchange = () => { S.cfg.agg = D.cfgAgg.value; renderChart(); };
      D.cfgTopN.onchange = () => { S.cfg.topN = +D.cfgTopN.value; renderChart(); };
      D.cfgSortBy.onchange = () => { S.cfg.sortBy = D.cfgSortBy.value; renderChart(); };
      D.cfgBin.oninput = () => { S.cfg.bin = +D.cfgBin.value; D.cfgBinVal.textContent = D.cfgBin.value; scheduleRenderChart(); };
      D.cfgTitle.oninput = debounce(() => { S.cfg.title = D.cfgTitle.value.trim(); renderChart(); }, 300);
      [['cfgSmooth', 'smooth'], ['cfgStack', 'stack'], ['cfgLabel', 'label'], ['cfgRotate', 'rotate'], ['cfgDataZoom', 'zoom']].forEach(pair => {
        D[pair[0]].onchange = () => { S.cfg[pair[1]] = D[pair[0]].checked; renderChart(); };
      });
      D.yList.onclick = e => {
        const cb = e.target.closest('[data-y]'); if (!cb) return;
        const i = +cb.dataset.y;
        const single = ['pie', 'ring', 'funnel'].indexOf(S.cfg.type) >= 0;
        if (cb.checked) { if (single) S.cfg.y = [i]; else { if (S.cfg.y.indexOf(i) < 0) S.cfg.y.push(i); if (S.cfg.y.length > 6) { S.cfg.y.shift(); toast('最多同时对比 6 个指标', { kind: 'warn', ms: 2000 }); } } }
        else S.cfg.y = S.cfg.y.filter(x => x !== i);
        syncYList(); renderChart();
      };
      D.btnRenderChart.onclick = () => renderChart();
      D.btnChartAuto.onclick = () => autoConfigChart();
      D.btnChartReset.onclick = () => {
        S.cfg = Object.assign(S.cfg, { type: 'bar', mode: 'agg', split: -1, agg: 'sum', topN: 20, sortBy: 'xasc', bin: 14, smooth: true, stack: false, label: false, rotate: true, zoom: false, title: '' });
        syncCfgUI(); autoConfigChart(true); toast('图表配置已重置', { kind: 'ok' });
      };
      D.btnSaveChart.onclick = saveChart;
      D.btnClearSaved.onclick = async () => {
        if (!S.saved.length) return;
        if (await confirmDlg({ title: '清除固定图表', message: '将移除 ' + S.saved.length + ' 张已固定的图表，报告也会同步更新。', okText: '清除', danger: true })) {
          S.saved = []; renderSaved(); buildReportIfActive(); toast('已清除全部固定图表', { kind: 'ok' });
        }
      };
      D.btnChartFull.onclick = toggleFull;
      D.btnChartPng.onclick = () => {
        if (!S.chart) { toast('请先生成图表', { kind: 'warn' }); return; }
        const url = S.chart.getDataURL({ pixelRatio: 2, backgroundColor: chartColors().bg });
        const a = document.createElement('a');
        a.href = url; a.download = stampName(D.chartTitle.textContent.replace(/[\\/:*?"<>|]/g, '_').slice(0, 40) || 'chart', 'png');
        a.click(); toast('已导出高清 PNG', { kind: 'ok' });
      };
      D.btnChartJson.onclick = async () => {
        if (!S.chartOpt) { toast('请先生成图表', { kind: 'warn' }); return; }
        const ok = await copyText(JSON.stringify(S.chartOpt, null, 2));
        toast(ok ? 'ECharts 配置 JSON 已复制到剪贴板' : '复制失败，请手动选择', { kind: ok ? 'ok' : 'err' });
      };
    }

    function toggleFull() {
      S.full = !S.full;
      D.chartBox.classList.toggle('fs', S.full);
      D.btnChartFull.innerHTML = '<svg class="ic ic-sm"><use href="#' + (S.full ? 'i-shrink' : 'i-expand') + '"/></svg>';
      document.body.style.overflow = S.full ? 'hidden' : '';
      setTimeout(() => { if (S.chart) S.chart.resize(); }, 80);
    }

    function saveChart() {
      if (!S.chart || !S.chartOpt) { toast('请先生成图表再固定', { kind: 'warn' }); return; }
      let img = '';
      try { img = S.chart.getDataURL({ pixelRatio: 1.3, backgroundColor: chartColors().bg }); } catch (e) { }
      S.saved.unshift({
        id: uid(), title: D.chartTitle.textContent, img,
        meta: (D.chartFoot.textContent || '').trim().slice(0, 160),
        type: S.cfg.type, at: Date.now()
      });
      if (S.saved.length > 14) S.saved.pop();
      renderSaved(); buildReportIfActive();
      toast('已固定到图库，并自动加入报告', { kind: 'ok', title: '图表已保存', action: '查看', onAction: () => { go('report'); } });
    }

    function renderSaved() {
      D.savedCount.textContent = String(S.saved.length);
      if (!S.saved.length) {
        D.savedGrid.innerHTML = emptyBox('i-pin', '还没有固定图表', '在上方配置好图表后点击「固定」，即可在这里管理，并自动汇总到报告中。');
        return;
      }
      D.savedGrid.innerHTML = '<div class="thumb-grid">' + S.saved.map(s =>
        '<div class="thumb" data-id="' + s.id + '">' +
        (s.img ? '<img src="' + s.img + '" alt="' + esc(s.title) + '">' : '<div style="height:168px" class="center dim tiny">无预览</div>') +
        '<div class="thumb-b"><span class="grow truncate tiny b" title="' + esc(s.title) + '">' + esc(s.title) + '</span>' +
        '<button class="btn btn-ghost btn-icon btn-sm" data-a="view" aria-label="放大" data-tip="放大"><svg class="ic ic-sm"><use href="#i-expand"/></svg></button>' +
        '<button class="btn btn-ghost btn-icon btn-sm" data-a="dl" aria-label="下载" data-tip="下载 PNG"><svg class="ic ic-sm"><use href="#i-download"/></svg></button>' +
        '<button class="btn btn-ghost btn-icon btn-sm" data-a="del" aria-label="移除" data-tip="移除"><svg class="ic ic-sm"><use href="#i-trash"/></svg></button>' +
        '</div><div class="tiny dim truncate" style="padding:0 12px 10px">' + esc(s.meta || '') + '</div></div>').join('') + '</div>';
      D.savedGrid.querySelectorAll('.thumb').forEach(card => {
        const id = card.dataset.id;
        const s = S.saved.filter(x => x.id === id)[0];
        card.querySelector('[data-a="view"]').onclick = () => openModal({
          title: s.title, icon: 'i-chart', size: 'wide',
          body: '<img src="' + s.img + '" style="width:100%;border-radius:12px;border:1px solid var(--border)" alt="' + esc(s.title) + '">',
          footer: '<button class="btn" data-close>关闭</button><button class="btn btn-primary" data-dl>下载 PNG</button>',
          onMount(root, api) { root.querySelector('[data-dl]').onclick = () => { const a = document.createElement('a'); a.href = s.img; a.download = stampName(s.title.slice(0, 40), 'png'); a.click(); api.close(); }; }
        });
        card.querySelector('[data-a="dl"]').onclick = () => { const a = document.createElement('a'); a.href = s.img; a.download = stampName(s.title.slice(0, 40), 'png'); a.click(); };
        card.querySelector('[data-a="del"]').onclick = () => { S.saved = S.saved.filter(x => x.id !== id); renderSaved(); buildReportIfActive(); toast('已移除该图表', { kind: 'ok', ms: 1800 }); };
      });
    }

/* =========================================================
       透视分析
       ========================================================= */
    function bindPivotUI() {
      D.pvRowList.onclick = e => {
        const cb = e.target.closest('[data-pvr]'); if (!cb) return;
        const i = +cb.dataset.pvr;
        if (cb.checked) { if (S.pv.rows.indexOf(i) < 0) S.pv.rows.push(i); if (S.pv.rows.length > 3) { S.pv.rows.shift(); syncPvRows(); toast('最多支持 3 级行维度', { kind: 'warn', ms: 2000 }); } }
        else S.pv.rows = S.pv.rows.filter(x => x !== i);
        if (S.pv.rows.indexOf(S.pv.col) >= 0) { S.pv.col = -1; D.pvCol.value = '-1'; }
        renderPivot();
      };
      D.pvCol.onchange = () => {
        S.pv.col = +D.pvCol.value;
        if (S.pv.rows.indexOf(S.pv.col) >= 0) { toast('列维度不能与行维度重复', { kind: 'warn' }); S.pv.col = -1; D.pvCol.value = '-1'; }
        renderPivot();
      };
      D.pvVal.onchange = () => { S.pv.val = +D.pvVal.value; renderPivot(); };
      D.pvAgg.onchange = () => { S.pv.agg = D.pvAgg.value; renderPivot(); };
      [['pvTotals', 'totals'], ['pvHeat', 'heat'], ['pvPercent', 'pct'], ['pvSortDesc', 'sortDesc']].forEach(p => {
        D[p[0]].onchange = () => { S.pv[p[1]] = D[p[0]].checked; renderPivot(); };
      });
      D.pvScope.onclick = e => {
        const b = e.target.closest('[data-s]'); if (!b) return;
        S.pv.scope = b.dataset.s;
        $$('#pvScope button').forEach(x => x.classList.toggle('on', x === b));
        renderPivot();
      };
      D.btnPivotRun.onclick = () => { renderPivot(); toast('透视表已更新', { kind: 'ok', ms: 1600 }); };
      D.btnPivotCopy.onclick = async () => {
        if (!S.pv.last) { toast('请先生成透视表', { kind: 'warn' }); return; }
        const ok = await copyText(toTSV(S.pv.last.head, S.pv.last.matrix));
        toast(ok ? '透视结果已复制（可直接粘贴到 Excel）' : '复制失败', { kind: ok ? 'ok' : 'err' });
      };
      D.btnPivotExport.onclick = () => {
        if (!S.pv.last) { toast('请先生成透视表', { kind: 'warn' }); return; }
        if (!window.XLSX) { downloadText('\uFEFF' + toCSV(S.pv.last.head, S.pv.last.matrix), stampName('透视表', 'csv'), 'text/csv'); toast('已导出 CSV', { kind: 'ok' }); return; }
        const ws = XLSX.utils.aoa_to_sheet([S.pv.last.head].concat(S.pv.last.matrix));
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '透视表');
        XLSX.writeFile(wb, stampName('透视表', 'xlsx'));
        toast('已导出 Excel 透视表', { kind: 'ok' });
      };
    }
    function syncPvRows() { D.pvRowList.querySelectorAll('[data-pvr]').forEach(cb => cb.checked = S.pv.rows.indexOf(+cb.dataset.pvr) >= 0); }

    function renderPivot() {
      if (!S.rows.length) {
        D.pvWrap.innerHTML = ''; D.pvInfo.textContent = '—';
        D.pvEmpty.innerHTML = emptyBox('i-grid', '还没有数据', '导入数据后即可进行交叉透视分析。');
        return;
      }
      D.pvEmpty.innerHTML = '';
      if (!S.pv.rows.length) {
        D.pvWrap.innerHTML = emptyBox('i-grid', '请选择行维度', '在左侧勾选一个或多个行维度（最多 3 级），即可生成交叉汇总表。');
        D.pvInfo.textContent = '未配置';
        S.pv.last = null;
        return;
      }
      const cfg = S.pv, rows = S.rows;
      const idx = cfg.scope === 'filtered' ? S.view.idx : rows.map((r, i) => i);
      const valIdx = cfg.val >= 0 && S.cols[cfg.val] ? cfg.val : -1;
      const agg = valIdx < 0 ? 'countAll' : cfg.agg;
      const colIdx = cfg.col >= 0 && S.cols[cfg.col] && cfg.rows.indexOf(cfg.col) < 0 ? cfg.col : -1;

      const rowMap = new Map(), colCount = new Map();
      const labelOf = (ri, c) => { const v = fmtVal(rows[ri][c], S.cols[c].type); return String(v).trim() === '' ? '（空）' : shortLabel(v, 22); };
      for (const ri of idx) {
        const labels = cfg.rows.map(c => labelOf(ri, c));
        const rk = labels.join('\u0002');
        let e = rowMap.get(rk);
        if (!e) { e = { labels, cells: new Map() }; rowMap.set(rk, e); }
        const ck = colIdx >= 0 ? labelOf(ri, colIdx) : '__all__';
        colCount.set(ck, (colCount.get(ck) || 0) + 1);
        if (!e.cells.has(ck)) e.cells.set(ck, []);
        e.cells.get(ck).push(ri);
      }
      let colKeys = colIdx >= 0 ? Array.from(colCount.entries()).sort((a, b) => b[1] - a[1]).map(e => e[0]) : ['__all__'];
      let otherCols = null;
      if (colKeys.length > 24) { otherCols = colKeys.slice(24); colKeys = colKeys.slice(0, 24); }
      if (rowMap.size > 4000) { toast('透视行过多，仅显示前 4000 组', { kind: 'warn' }); }

      const entries = Array.from(rowMap.values());
      const rowTotal = e => aggValues([].concat.apply([], Array.from(e.cells.values())), rows, valIdx, agg);
      entries.forEach(e => { e.total = rowTotal(e); });
      if (cfg.sortDesc) entries.sort((a, b) => (b.total || 0) - (a.total || 0));
      const limited = entries.slice(0, 4000);

      const cellVal = (e, ck) => {
        let list = e.cells.get(ck) || [];
        if (!list.length && otherCols && ck === '__other__') list = [];
        return aggValues(list, rows, valIdx, agg);
      };
      const colTotal = ck => {
        const list = [];
        limited.forEach(e => { const l = e.cells.get(ck); if (l) for (let i = 0; i < l.length; i++) list.push(l[i]); });
        return aggValues(list, rows, valIdx, agg);
      };
      const allList = [].concat.apply([], limited.map(e => Array.from(e.cells.values())).map(a => a));
      const grand = aggValues(idx.slice(), rows, valIdx, agg);

      let maxAbs = 0;
      limited.forEach(e => colKeys.forEach(ck => { const v = Math.abs(cellVal(e, ck)); if (v > maxAbs) maxAbs = v; }));

      const lv = cfg.rows.length;
      const spans = [];
      for (let L = 0; L < lv; L++) {
        let i = 0;
        while (i < limited.length) {
          let j = i + 1;
          while (j < limited.length && limited[j].labels.slice(0, L + 1).join('\u0002') === limited[i].labels.slice(0, L + 1).join('\u0002')) j++;
          spans.push({ L, i, n: j - i });
          i = j;
        }
      }
      const spanAt = (L, i) => { const s = spans.filter(x => x.L === L && x.i === i)[0]; return s ? s.n : 0; };

      const showCols = colKeys;
      const headRow1 = '<tr>' + cfg.rows.map((c, L) => '<th class="rowh" rowspan="2">' + esc(S.cols[c].name) + '</th>').join('') +
        (colIdx >= 0 ? '<th colspan="' + (showCols.length + (otherCols ? 1 : 0)) + '" style="text-align:center">' + esc(S.cols[colIdx].name) + '</th>' : '<th rowspan="2">' + esc(valIdx >= 0 ? AGG_NAMES[agg] + '(' + S.cols[valIdx].name + ')' : '行数') + '</th>') +
        (cfg.totals ? '<th class="rowh total" rowspan="2">总计</th>' : '') + '</tr>';
      const headRow2 = colIdx >= 0 ? '<tr>' + showCols.map(ck => '<th>' + esc(ck) + '</th>').join('') + (otherCols ? '<th>其他 ' + otherCols.length + ' 项</th>' : '') + '</tr>' : '';

      const heatBg = v => {
        if (!cfg.heat || !maxAbs) return '';
        const a = clamp(Math.abs(v) / maxAbs, 0, 1);
        return 'background:color-mix(in srgb, var(--brand) ' + (a * 46).toFixed(1) + '%, transparent);' + (a > .62 ? 'color:#fff;' : '');
      };
      let body = '';
      limited.forEach((e, i) => {
        body += '<tr>';
        for (let L = 0; L < lv; L++) {
          const n = spanAt(L, i);
          if (n > 0) body += '<td class="rowh" rowspan="' + n + '">' + esc(e.labels[L]) + (L === lv - 1 && n > 1 ? '' : '') + '</td>';
        }
        if (colIdx >= 0) {
          showCols.forEach(ck => {
            const v = cellVal(e, ck);
            body += '<td style="' + heatBg(v) + '">' + fmtNum(v, Math.abs(v) < 100 && v % 1 ? 2 : 0) +
              (cfg.pct && grand ? '<div class="xs" style="opacity:.62">' + (v / grand * 100).toFixed(1) + '%</div>' : '') + '</td>';
          });
          if (otherCols) {
            let v = 0; otherCols.forEach(ck => { v += cellVal(e, ck); });
            body += '<td class="dim">' + (agg === 'sum' || agg === 'count' || agg === 'countAll' ? fmtNum(v) : '—') + '</td>';
          }
          if (cfg.totals) body += '<td class="total">' + fmtNum(e.total, Math.abs(e.total) < 100 && e.total % 1 ? 2 : 0) + '</td>';
        } else {
          body += '<td style="' + heatBg(e.total) + '">' + fmtNum(e.total, Math.abs(e.total) < 100 && e.total % 1 ? 2 : 0) +
            (cfg.pct && grand ? '<div class="xs" style="opacity:.62">' + (e.total / grand * 100).toFixed(1) + '%</div>' : '') + '</td>';
        }
        body += '</tr>';
      });
      if (cfg.totals && limited.length) {
        body += '<tr class="total">';
        body += '<td class="rowh total" colspan="' + lv + '">总计（' + fmtNum(limited.length) + ' 组）</td>';
        if (colIdx >= 0) {
          showCols.forEach(ck => { body += '<td class="total">' + fmtNum(colTotal(ck)) + '</td>'; });
          if (otherCols) body += '<td class="total">—</td>';
        }
        body += '<td class="total">' + fmtNum(grand) + '</td></tr>';
      }

      D.pvWrap.innerHTML = '<table class="pvt"><thead>' + headRow1 + headRow2 + '</thead><tbody>' + body + '</tbody></table>';
      D.pvTitle.textContent = cfg.rows.map(c => S.cols[c].name).join(' / ') + ' × ' + (colIdx >= 0 ? S.cols[colIdx].name : AGG_NAMES[agg]) ;
      D.pvInfo.textContent = fmtNum(limited.length) + ' 行 × ' + (colIdx >= 0 ? showCols.length + (otherCols ? 1 : 0) : 1) + ' 列 · ' + fmtNum(idx.length) + ' 条记录 · ' + AGG_NAMES[agg];

      // 供复制 / 导出 / 报告使用的扁平矩阵
      const head = cfg.rows.map(c => S.cols[c].name).concat(colIdx >= 0 ? showCols.concat(otherCols ? ['其他'] : []) : [AGG_NAMES[agg] + '(' + (valIdx >= 0 ? S.cols[valIdx].name : '行数') + ')']).concat(cfg.totals ? ['总计'] : []);
      const matrix = limited.map(e => {
        const base = e.labels.slice();
        if (colIdx >= 0) {
          showCols.forEach(ck => base.push(round2(cellVal(e, ck))));
          if (otherCols) base.push(round2(otherCols.reduce((s, ck) => s + cellVal(e, ck), 0)));
        } else base.push(round2(e.total));
        if (cfg.totals) base.push(round2(e.total));
        return base;
      });
      S.pv.last = { head, matrix, title: D.pvTitle.textContent, info: D.pvInfo.textContent };
    }
    function round2(v) { return typeof v === 'number' ? +v.toFixed(4) : v; }

    /* =========================================================
       数据清洗
       ========================================================= */
    function canSnapshot() { return S.rows.length * S.cols.length <= 400000; }
    function snapshot(label) {
      if (!canSnapshot()) { toast('数据量较大，本次操作不记录历史（仍可用「恢复原始数据」回退）', { kind: 'warn', ms: 4200 }); return false; }
      S.hist.push({ label, at: Date.now(), cols: S.cols.map(c => ({ ...c })), rows: S.rows.map(r => r.slice()) });
      if (S.hist.length > 16) S.hist.shift();
      S.future = [];
      renderHist(); updateSourceCard();
      return true;
    }
    function undo() {
      if (!S.hist.length) { toast('没有可撤销的操作', { kind: 'info', ms: 1800 }); return; }
      const h = S.hist.pop();
      S.future.push({ label: h.label, at: Date.now(), cols: S.cols.map(c => ({ ...c })), rows: S.rows.map(r => r.slice()) });
      S.cols = h.cols; S.rows = h.rows;
      afterMutation('已撤销：' + h.label);
    }
    function redo() {
      if (!S.future.length) { toast('没有可重做的操作', { kind: 'info', ms: 1800 }); return; }
      const h = S.future.pop();
      S.hist.push({ label: h.label, at: Date.now(), cols: S.cols.map(c => ({ ...c })), rows: S.rows.map(r => r.slice()) });
      S.cols = h.cols; S.rows = h.rows;
      afterMutation('已重做：' + h.label);
    }
    function afterMutation(msg) {
      bumpVer(); S.dirty = true; S.sel.clear();
      S.view.filters = {}; S.view.search = ''; S.view.tokens = []; D.tblSearch.value = '';
      S.view.hidden = S.view.hidden.filter(i => i < S.cols.length);
      if (S.view.sort.c >= S.cols.length) S.view.sort = { c: -1, d: 0 };
      reprofile(); computeIdx();
      S.view.page = 1;
      renderTable(); renderOverview(); updateSourceCard(); fillColumnSelects();
      syncCfgUI(); renderChart(); renderPivotUI(); renderCleanUI(); renderHist();
      buildReportIfActive(); maybeAutosave();
      if (msg) toast(msg, { kind: 'ok' });
    }

    function renderHist() {
      D.navHistCount.textContent = String(S.hist.length);
      D.histCount.textContent = String(S.hist.length);
      if (!S.hist.length) {
        D.histList.innerHTML = '<div class="empty" style="padding:26px"><div class="empty-ic" style="width:52px;height:52px"><svg class="ic"><use href="#i-history"/></svg></div><p class="tiny">还没有清洗操作。<br>每一步修改都会记录在这里，可随时撤销。</p></div>';
        return;
      }
      D.histList.innerHTML = S.hist.slice().reverse().map((h, k) => {
        const n = S.hist.length - k;
        return '<div class="hist-item"><span class="hist-n">' + n + '</span><div class="grow" style="min-width:0">' +
          '<div class="tiny b truncate">' + esc(h.label) + '</div>' +
          '<div class="xs dim">' + new Date(h.at).toLocaleTimeString('zh-CN') + ' · ' + fmtNum(h.rows.length) + ' 行 × ' + h.cols.length + ' 列</div></div>' +
          (k === 0 ? '<button class="btn btn-xs btn-soft" data-undo>撤销这一步</button>' : '') + '</div>';
      }).join('') + (S.future.length ? '<div class="hist-item" style="opacity:.7"><span class="hist-n" style="background:var(--soft);color:var(--text-3)">' + S.future.length + '</span><div class="grow tiny dim">可重做的操作</div><button class="btn btn-xs" data-redo>重做</button></div>' : '');
      const u = D.histList.querySelector('[data-undo]'); if (u) u.onclick = undo;
      const r = D.histList.querySelector('[data-redo]'); if (r) r.onclick = redo;
    }

    function renderCleanUI() {
      if (!S.rows.length) {
        D.cleanRowInfo.textContent = '—';
        D.cleanDiff.innerHTML = emptyBox('i-sparkles', '暂无数据', '导入数据后可在此进行清洗与转换。');
        renderHist();
        return;
      }
      D.cleanRowInfo.textContent = fmtNum(S.rows.length) + ' 行 × ' + S.cols.length + ' 列';
      const o = S.origin, miss = S.profile ? S.profile.reduce((s, p) => s + p.blank, 0) : 0;
      const cells = S.rows.length * S.cols.length;
      const oMiss = o ? countMissing(o.rows, o.cols.length) : 0;
      const d = (a, b, unit, good) => {
        const diff = a - b;
        if (!diff) return '<span class="badge">无变化</span>';
        const positive = diff > 0;
        const tone = (positive === good) ? 'badge-ok' : 'badge-warn';
        return '<span class="badge ' + tone + '">' + (positive ? '+' : '') + fmtNum(diff) + (unit || '') + '</span>';
      };
      D.cleanDiff.innerHTML =
        '<div class="stat-line"><span class="muted">数据行数</span><span>' + fmtNum(o ? o.rows.length : S.rows.length) + ' → <b>' + fmtNum(S.rows.length) + '</b> ' + d(S.rows.length, o ? o.rows.length : S.rows.length, ' 行', true) + '</span></div>' +
        '<div class="stat-line"><span class="muted">数据列数</span><span>' + (o ? o.cols.length : S.cols.length) + ' → <b>' + S.cols.length + '</b> ' + d(S.cols.length, o ? o.cols.length : S.cols.length, ' 列', true) + '</span></div>' +
        '<div class="stat-line"><span class="muted">缺失单元格</span><span>' + fmtNum(oMiss) + ' → <b>' + fmtNum(miss) + '</b> ' + d(miss, oMiss, '', false) + '</span></div>' +
        '<div class="stat-line"><span class="muted">数据完整度</span><span><b>' + (cells ? (100 - miss / cells * 100).toFixed(1) : '100.0') + '%</b></span></div>' +
        '<div class="stat-line"><span class="muted">重复行</span><span><b>' + fmtNum(countDupRows()) + '</b></span></div>' +
        '<div class="stat-line"><span class="muted">已记录步骤</span><span><b>' + S.hist.length + '</b> / 16</span></div>' +
        '<div class="hint" style="margin-top:8px">' + (S.dirty ? '数据已被修改，导出时会包含这些改动。' : '当前数据与导入时一致。') + '</div>';
      renderHist();
      D.btnUndo.disabled = !S.hist.length;
    }
    function countMissing(rows, nCols) {
      let m = 0;
      for (let i = 0; i < rows.length; i++) for (let c = 0; c < nCols; c++) if (isBlank(rows[i][c])) m++;
      return m;
    }

/* =========================================================
       数据清洗：具体操作
       ========================================================= */
    const CLEAN_LABEL = {
      dropEmptyRows: '删除完全空行', dropDupRows: '删除重复行', dropRowsMissingAny: '删除含缺失值的行',
      trimAll: '去除文本首尾空格', collapseSpace: '合并多余空白', shuffle: '随机打乱行序'
    };

    function runClean(act) {
      if (!S.rows.length) { toast('请先导入数据', { kind: 'warn' }); return; }
      const before = S.rows.length;
      let changed = 0;
      const pushed = snapshot(CLEAN_LABEL[act] || act);
      if (act === 'dropEmptyRows') S.rows = S.rows.filter(r => !r.every(isBlank));
      else if (act === 'dropDupRows') {
        const seen = new Set(), out = [];
        for (const r of S.rows) {
          const k = r.map(v => v instanceof Date ? 'd' + v.getTime() : String(v)).join('\u0001');
          if (seen.has(k)) continue;
          seen.add(k); out.push(r);
        }
        S.rows = out;
      }
      else if (act === 'dropRowsMissingAny') S.rows = S.rows.filter(r => !r.some(isBlank));
      else if (act === 'trimAll') {
        for (const r of S.rows) for (let c = 0; c < r.length; c++) {
          if (typeof r[c] === 'string') { const t = r[c].trim(); if (t !== r[c]) { r[c] = t; changed++; } }
        }
      }
      else if (act === 'collapseSpace') {
        for (const r of S.rows) for (let c = 0; c < r.length; c++) {
          if (typeof r[c] === 'string') { const t = r[c].replace(/\s+/g, ' ').trim(); if (t !== r[c]) { r[c] = t; changed++; } }
        }
      }
      else if (act === 'shuffle') {
        for (let i = S.rows.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const t = S.rows[i]; S.rows[i] = S.rows[j]; S.rows[j] = t; }
        changed = S.rows.length;
      }
      const dn = before - S.rows.length;
      let msg = CLEAN_LABEL[act] + '：';
      if (dn > 0) msg += '删除 ' + fmtNum(dn) + ' 行，剩余 ' + fmtNum(S.rows.length) + ' 行';
      else if (changed) msg += '修改 ' + fmtNum(changed) + ' ' + (act === 'shuffle' ? '行位置' : '个单元格');
      else msg += '没有需要处理的内容';
      afterMutation(msg);
      if (dn === 0 && !changed && pushed) { S.hist.pop(); renderHist(); }
    }

    function fillColumn(ci, mode, text) {
      if (ci < 0 || !S.cols[ci]) return;
      const name = S.cols[ci].name;
      const blanks = S.rows.reduce((s, r) => s + (isBlank(r[ci]) ? 1 : 0), 0);
      if (!blanks) { toast('「' + name + '」没有缺失值', { kind: 'info' }); return; }
      const pushed = snapshot('填充缺失值：' + name + '（' + ({ zero: '填 0', mean: '均值', median: '中位数', mode: '众数', ffill: '上一行值', text: '指定文本', drop: '删除行' })[mode] + '）');

      if (mode === 'drop') {
        S.rows = S.rows.filter(r => !isBlank(r[ci]));
        afterMutation('已删除「' + name + '」为空的 ' + blanks + ' 行');
        return;
      }
      let fill;
      if (mode === 'mean' || mode === 'median') {
        const p = S.profile[ci];
        if (!p || p.type !== 'number' || !p.count) { toast('该列不是数值列，无法用均值 / 中位数填充', { kind: 'warn' }); if (pushed) { S.hist.pop(); renderHist(); } return; }
        fill = mode === 'mean' ? +p.mean.toFixed(4) : +p.median.toFixed(4);
      } else if (mode === 'mode') {
        const f = new Map();
        for (const r of S.rows) { const v = r[ci]; if (isBlank(v)) continue; const k = v instanceof Date ? v.getTime() : String(v); f.set(k, (f.get(k) || 0) + 1); }
        if (!f.size) { toast('该列没有可用于填充的众数', { kind: 'warn' }); if (pushed) { S.hist.pop(); renderHist(); } return; }
        const top = Array.from(f.entries()).sort((a, b) => b[1] - a[1])[0][0];
        const sample = S.rows.map(r => r[ci]).filter(v => !isBlank(v)).filter(v => String(v instanceof Date ? v.getTime() : v) === String(top))[0];
        fill = sample;
      } else if (mode === 'zero') fill = S.cols[ci].type === 'number' ? 0 : '0';
      else if (mode === 'text') fill = text === undefined || text === null || text === '' ? '未知' : text;

      let n = 0;
      if (mode === 'ffill') {
        let last = null;
        for (const r of S.rows) { if (isBlank(r[ci])) { if (last !== null) { r[ci] = last; n++; } } else last = r[ci]; }
      } else {
        for (const r of S.rows) if (isBlank(r[ci])) { r[ci] = fill; n++; }
      }
      afterMutation('已填充「' + name + '」的 ' + fmtNum(n) + ' 个缺失值' + (n < blanks ? '（其余 ' + (blanks - n) + ' 个之前没有可参考的值）' : ''));
    }

    function convertType(ci, type) {
      if (ci < 0 || !S.cols[ci]) return;
      const name = S.cols[ci].name;
      let fail = 0, n = 0;
      const preview = [];
      for (let i = 0; i < S.rows.length; i++) {
        const v = S.rows[i][ci];
        if (isBlank(v)) continue;
        let nv;
        if (type === 'number') { const x = toNum(v); if (isNaN(x)) { fail++; nv = v; } else nv = x; }
        else if (type === 'date') { const d = toDate(v); if (!d) { fail++; nv = v; } else nv = d; }
        else nv = String(fmtVal(v, S.cols[ci].type));
        if (nv !== v) n++;
        if (preview.length < 3 && nv !== v) preview.push(String(fmtVal(v, S.cols[ci].type)) + ' → ' + String(fmtVal(nv, type)));
      }
      snapshot('转换列类型：' + name + ' → ' + TYPE_LABEL[type]);
      for (let i = 0; i < S.rows.length; i++) {
        const v = S.rows[i][ci];
        if (isBlank(v)) continue;
        if (type === 'number') { const x = toNum(v); if (!isNaN(x)) S.rows[i][ci] = x; }
        else if (type === 'date') { const d = toDate(v); if (d) S.rows[i][ci] = d; }
        else S.rows[i][ci] = String(fmtVal(v, S.cols[ci].type));
      }
      S.cols[ci].type = type;
      afterMutation('「' + name + '」已转为' + TYPE_LABEL[type] + '，成功 ' + fmtNum(n) + ' 个' + (fail ? '，' + fail + ' 个无法解析已保留原值' : '') + (preview.length ? '（例：' + preview[0] + '）' : ''));
    }

    function remapAfterDropCol(ci) {
      S.view.hidden = S.view.hidden.filter(i => i !== ci).map(i => i > ci ? i - 1 : i);
      if (S.view.sort.c === ci) S.view.sort = { c: -1, d: 0 };
      else if (S.view.sort.c > ci) S.view.sort.c--;
      const nf = {};
      Object.keys(S.view.filters).forEach(k => { const i = +k; if (i === ci) return; nf[i > ci ? i - 1 : i] = S.view.filters[k]; });
      S.view.filters = nf;
      S.cfg.x = S.cfg.x === ci ? -1 : (S.cfg.x > ci ? S.cfg.x - 1 : S.cfg.x);
      S.cfg.split = S.cfg.split === ci ? -1 : (S.cfg.split > ci ? S.cfg.split - 1 : S.cfg.split);
      S.cfg.y = S.cfg.y.filter(i => i !== ci).map(i => i > ci ? i - 1 : i);
      S.pv.rows = S.pv.rows.filter(i => i !== ci).map(i => i > ci ? i - 1 : i);
      S.pv.col = S.pv.col === ci ? -1 : (S.pv.col > ci ? S.pv.col - 1 : S.pv.col);
      S.pv.val = S.pv.val === ci ? -1 : (S.pv.val > ci ? S.pv.val - 1 : S.pv.val);
    }

    async function dropCol(ci) {
      if (ci < 0 || !S.cols[ci]) return;
      if (S.cols.length <= 1) { toast('至少需要保留一列', { kind: 'warn' }); return; }
      const ok = await confirmDlg({ title: '删除列', message: '确定删除列「' + S.cols[ci].name + '」吗？', detail: '该列的 ' + fmtNum(S.rows.length) + ' 个值将被移除，可通过撤销恢复。', okText: '删除列', danger: true });
      if (!ok) return;
      snapshot('删除列：' + S.cols[ci].name);
      const name = S.cols[ci].name;
      S.cols.splice(ci, 1);
      S.rows.forEach(r => r.splice(ci, 1));
      remapAfterDropCol(ci);
      closePop();
      const dr = D.drawerRoot.querySelector('.drawer'); if (dr) D.drawerRoot.innerHTML = '';
      afterMutation('已删除列「' + name + '」');
    }

    async function renameCol(ci) {
      if (ci < 0 || !S.cols[ci]) return;
      const nv = await promptDlg({ title: '重命名列', label: '新的列名', value: S.cols[ci].name, required: true, icon: 'i-pencil', hint: '列名会同步更新到图表、透视表与导出文件中' });
      if (nv === null) return;
      const name = nv.trim().slice(0, 60);
      if (!name) return;
      if (S.cols.some((c, i) => i !== ci && c.name.toLowerCase() === name.toLowerCase())) { toast('已存在同名列，请换一个名字', { kind: 'warn' }); return; }
      snapshot('重命名列：' + S.cols[ci].name + ' → ' + name);
      const old = S.cols[ci].name;
      S.cols[ci].name = name;
      afterMutation('列名已由「' + old + '」改为「' + name + '」');
    }

    async function splitColumn(ci, sep) {
      if (ci < 0 || !S.cols[ci]) return;
      sep = sep || '-';
      let maxParts = 1;
      for (const r of S.rows) { const v = r[ci]; if (isBlank(v)) continue; maxParts = Math.max(maxParts, String(v).split(sep).length); }
      if (maxParts < 2) { toast('用「' + sep + '」拆分后只有 1 段，请检查分隔符', { kind: 'warn' }); return; }
      maxParts = Math.min(maxParts, 10);
      const ok = await confirmDlg({ title: '拆分列', message: '将「' + S.cols[ci].name + '」按「' + sep + '」拆成 ' + maxParts + ' 列。', detail: '原列会被替换为 ' + S.cols[ci].name + '_1 … ' + S.cols[ci].name + '_' + maxParts + '，可通过撤销恢复。', okText: '开始拆分' });
      if (!ok) return;
      snapshot('拆分列：' + S.cols[ci].name + ' → ' + maxParts + ' 列');
      const base = S.cols[ci].name;
      const newCols = [];
      for (let k = 0; k < maxParts; k++) newCols.push({ name: base + '_' + (k + 1), type: 'text' });
      const rows = S.rows.map(r => {
        const v = r[ci]; const parts = isBlank(v) ? [] : String(v).split(sep);
        const out = r.slice(0, ci);
        for (let k = 0; k < maxParts; k++) out.push(parts[k] === undefined ? '' : String(parts[k]).trim());
        return out.concat(r.slice(ci + 1));
      });
      S.rows = rows;
      S.cols.splice.apply(S.cols, [ci, 1].concat(newCols));
      remapAfterDropCol(ci);
      for (let k = 0; k < maxParts; k++) S.view.hidden = S.view.hidden.filter(i => i !== ci + k);
      newCols.forEach((c, k) => { c.type = inferType(S.rows.map(r => r[ci + k])); });
      afterMutation('「' + base + '」已拆分为 ' + maxParts + ' 列');
    }

    async function dedupeBy(ci) {
      if (ci < 0 || !S.cols[ci]) return;
      const seen = new Set(); let dup = 0;
      for (const r of S.rows) { const k = valKey(r[ci], S.cols[ci].type); if (seen.has(k)) dup++; else seen.add(k); }
      if (!dup) { toast('按「' + S.cols[ci].name + '」没有发现重复值', { kind: 'info' }); return; }
      const ok = await confirmDlg({ title: '按列去重', message: '「' + S.cols[ci].name + '」有 ' + dup + ' 行重复，将保留每组第一次出现的行。', okText: '删除 ' + dup + ' 行', danger: true });
      if (!ok) return;
      snapshot('按「' + S.cols[ci].name + '」去重');
      const s2 = new Set(), out = [];
      for (const r of S.rows) { const k = valKey(r[ci], S.cols[ci].type); if (s2.has(k)) continue; s2.add(k); out.push(r); }
      S.rows = out;
      afterMutation('已按「' + S.cols[ci].name + '」去重，删除 ' + dup + ' 行');
    }

    /* ---------- 计算列 ---------- */
    const CALC_F = {
      IF: (c, a, b) => (c ? a : b),
      AND: function () { for (let i = 0; i < arguments.length; i++) if (!arguments[i]) return false; return true; },
      OR: function () { for (let i = 0; i < arguments.length; i++) if (arguments[i]) return true; return false; },
      NOT: c => !c,
      ABS: x => Math.abs(+x || 0), SQRT: x => Math.sqrt(+x || 0), POW: (x, y) => Math.pow(+x || 0, +y || 0),
      ROUND: (x, n) => { const p = Math.pow(10, +n || 0); return Math.round((+x || 0) * p) / p; },
      FLOOR: x => Math.floor(+x || 0), CEIL: x => Math.ceil(+x || 0),
      MIN: function () { let m = Infinity; for (let i = 0; i < arguments.length; i++) { const v = +arguments[i]; if (!isNaN(v) && v < m) m = v; } return m === Infinity ? 0 : m; },
      MAX: function () { let m = -Infinity; for (let i = 0; i < arguments.length; i++) { const v = +arguments[i]; if (!isNaN(v) && v > m) m = v; } return m === -Infinity ? 0 : m; },
      LEN: s => String(s === null || s === undefined ? '' : s).length,
      UPPER: s => String(s === null || s === undefined ? '' : s).toUpperCase(),
      LOWER: s => String(s === null || s === undefined ? '' : s).toLowerCase(),
      TRIM: s => String(s === null || s === undefined ? '' : s).trim(),
      CONCAT: function () { let s = ''; for (let i = 0; i < arguments.length; i++) s += arguments[i] === null || arguments[i] === undefined ? '' : String(arguments[i]); return s; },
      ISBLANK: v => isBlank(v),
      CONTAINS: (s, q) => String(s === null || s === undefined ? '' : s).indexOf(q) >= 0,
      NUM: v => { const n = toNum(v); return isNaN(n) ? 0 : n; },
      YEAR: v => { const d = toDate(v); return d ? d.getFullYear() : ''; },
      MONTH: v => { const d = toDate(v); return d ? d.getMonth() + 1 : ''; },
      DAY: v => { const d = toDate(v); return d ? d.getDate() : ''; },
      WEEKDAY: v => { const d = toDate(v); return d ? '周' + '日一二三四五六'[d.getDay()] : ''; }
    };

    function compileExpr(expr) {
      let src = String(expr || '').trim();
      if (!src) throw new Error('请输入表达式，例如 [单价] * [数量]');
      if (src.length > 500) throw new Error('表达式过长（上限 500 字符）');
      const refs = [];
      src = src.replace(/\[([^\]]+)\]/g, (m, nm0) => {
        const nm = nm0.trim();
        let i = S.cols.findIndex(c => c.name === nm);
        if (i < 0) i = S.cols.findIndex(c => c.name.toLowerCase() === nm.toLowerCase());
        if (i < 0) throw new Error('找不到列「' + nm + '」，请写成 [列名] 的形式');
        refs.push(i);
        return '__v(' + (refs.length - 1) + ')';
      });
      src = src.replace(/&&/g, '\u0001').replace(/&/g, '+').replace(/\u0001/g, '&&');
      src = src.replace(/\b(IF|AND|OR|NOT|ABS|SQRT|POW|ROUND|FLOOR|CEIL|MIN|MAX|LEN|UPPER|LOWER|TRIM|CONCAT|ISBLANK|CONTAINS|NUM|YEAR|MONTH|DAY|WEEKDAY)\s*\(/g, (m, f) => '__f.' + f + '(');
      if (/[;{}]|\bfunction\b|\breturn\b|=>|\bwindow\b|\bdocument\b|\beval\b|__proto__|\bthis\b|\bnew\b|\bimport\b|\bfetch\b|\[__/.test(src)) throw new Error('表达式包含不允许的语法');
      let fn;
      try { fn = new Function('__v', '__f', '"use strict"; return (' + src + ');'); }
      catch (e) { throw new Error('表达式语法错误：' + e.message); }
      return { fn, refs };
    }

    function evalCalc(comp, row) {
      const __v = k => {
        const ci = comp.refs[k], c = S.cols[ci], v = row[ci];
        if (c.type === 'number') { if (isBlank(v)) return 0; const n = toNum(v); return isNaN(n) ? 0 : n; }
        if (c.type === 'date') return v;
        if (isBlank(v)) return '';
        return v instanceof Date ? v : String(v);
      };
      return comp.fn(__v, CALC_F);
    }

    function calcPreview() {
      const name = D.calcName.value.trim();
      const expr = D.calcExpr.value.trim();
      D.calcPreview.innerHTML = '';
      if (!expr) { D.calcPreview.innerHTML = '<span class="dim">请先输入表达式</span>'; return null; }
      let comp;
      try { comp = compileExpr(expr); }
      catch (e) { D.calcPreview.innerHTML = '<span style="color:var(--err)">✗ ' + esc(e.message) + '</span>'; return null; }
      const out = [];
      let errs = 0;
      for (let i = 0; i < Math.min(5, S.rows.length); i++) {
        let v;
        try { v = evalCalc(comp, S.rows[i]); } catch (e) { errs++; v = '错误：' + e.message; }
        if (typeof v === 'number' && !isFinite(v)) { errs++; v = '无效数值'; }
        out.push('第 ' + (i + 1) + ' 行 → <b>' + esc(v instanceof Date ? fmtDate(v, true) : String(v)) + '</b>');
      }
      D.calcPreview.innerHTML = '<span style="color:var(--ok)">✓ 表达式有效</span> · 引用列：' +
        comp.refs.map(i => esc(S.cols[i].name)).join('、') + '<br>' + out.join('<br>') +
        (errs ? '<br><span style="color:var(--warn)">注意：有 ' + errs + ' 行计算异常</span>' : '');
      return { comp, name, errs };
    }

    async function applyCalc() {
      const r = calcPreview();
      if (!r) { toast('请先修正表达式', { kind: 'warn' }); return; }
      let name = r.name;
      if (!name) name = '计算列' + (S.cols.filter(c => /^计算列/.test(c.name)).length + 1);
      if (S.cols.some(c => c.name === name)) {
        const nn = await promptDlg({ title: '列名已存在', label: '请为新的计算列取名', value: name + '_2', required: true });
        if (nn === null) return;
        name = nn.trim();
      }
      const vals = []; let errs = 0;
      for (const row of S.rows) {
        let v;
        try { v = evalCalc(r.comp, row); } catch (e) { errs++; v = ''; }
        if (typeof v === 'number' && !isFinite(v)) { errs++; v = ''; }
        if (typeof v === 'boolean') v = v ? '是' : '否';
        vals.push(v === null || v === undefined ? '' : v);
      }
      snapshot('新增计算列：' + name);
      const ci = S.cols.length;
      S.cols.push({ name, type: 'text' });
      S.rows.forEach((row, i) => row.push(vals[i]));
      S.cols[ci].type = inferType(vals);
      afterMutation('已新增计算列「' + name + '」' + (S.cols[ci].type === 'number' ? '（数值）' : '') + (errs ? '，其中 ' + errs + ' 行计算失败已置空' : ''));
      D.calcName.value = ''; D.calcExpr.value = ''; D.calcPreview.innerHTML = '';
    }

    function calcHelp() {
      openModal({
        title: '计算列语法说明', icon: 'i-fx',
        body: '<p class="muted" style="font-size:13px">用 <span class="kbd">[列名]</span> 引用现有列，支持四则运算、括号与常用函数。文本拼接请用 <span class="kbd">&amp;</span> 或 <span class="kbd">CONCAT()</span>。</p>' +
          '<h4 class="b" style="margin:14px 0 6px;font-size:13.5px">示例</h4>' +
          '<div class="col gap-2">' +
          ['[单价] * [数量] * (1 - [折扣])', 'ROUND([销售额] / [数量], 2)', 'IF([是否退货] = "是", 0, [利润])', 'YEAR([订单日期]) & "-" & MONTH([订单日期])', 'UPPER(TRIM([商品]))', '[销售额] - [成本]', 'IF([满意度] >= 4.5, "好评", "一般")'].map(x =>
            '<code class="mono" style="padding:7px 10px;background:var(--soft);border-radius:8px;font-size:12.5px;cursor:pointer" data-fill="' + esc(x) + '">' + esc(x) + '</code>').join('') +
          '</div>' +
          '<h4 class="b" style="margin:16px 0 6px;font-size:13.5px">可用函数</h4>' +
          '<div class="tiny muted" style="line-height:1.9">' +
          '<b>逻辑</b> IF(条件, 真值, 假值) · AND() · OR() · NOT() · ISBLANK(x) · CONTAINS(x, "文本")<br>' +
          '<b>数学</b> ABS · ROUND(x, 位数) · SQRT · POW(x, y) · MIN · MAX · FLOOR · CEIL · NUM(x)<br>' +
          '<b>文本</b> LEN · UPPER · LOWER · TRIM · CONCAT(a, b, …)<br>' +
          '<b>日期</b> YEAR · MONTH · DAY · WEEKDAY</div>' +
          '<p class="hint" style="margin-top:12px">数值列中的空值按 0 参与计算；文本列中的空值按空字符串处理。计算在本地完成，不会上传数据。</p>',
        footer: '<button class="btn btn-primary" data-close>知道了</button>',
        onMount(root, api) {
          root.querySelectorAll('[data-fill]').forEach(c => c.onclick = () => { D.calcExpr.value = c.dataset.fill; api.close(); calcPreview(); D.calcExpr.focus(); });
        }
      });
    }

    function bindCleanUI() {
      document.querySelectorAll('#view-clean [data-act]').forEach(b => b.onclick = () => runClean(b.dataset.act));
      D.btnUndo.onclick = undo;
      D.btnResetData.onclick = async () => {
        if (!S.origin) { toast('没有可恢复的原始数据', { kind: 'warn' }); return; }
        const ok = await confirmDlg({ title: '恢复原始数据', message: '将丢弃全部清洗改动，回到刚导入时的状态。', detail: '共 ' + S.hist.length + ' 步操作会被撤销，此动作本身不可撤销。', okText: '恢复原始数据', danger: true });
        if (!ok) return;
        snapshot('恢复原始数据前的状态');
        S.cols = S.origin.cols.map(c => ({ ...c }));
        S.rows = S.origin.rows.map(r => r.slice());
        S.dirty = false;
        afterMutation('已恢复到导入时的原始数据');
      };
      D.fillCol.onchange = () => {
        const p = S.profile[+D.fillCol.value];
        D.fillPreview.textContent = p ? '该列缺失 ' + fmtNum(p.blank) + ' 个（' + (p.missing * 100).toFixed(1) + '%）' : '';
      };
      D.fillMode.onchange = () => { D.fillText.disabled = D.fillMode.value !== 'text'; };
      D.btnFill.onclick = () => fillColumn(+D.fillCol.value, D.fillMode.value, D.fillText.value);
      D.btnFillAllNumeric.onclick = async () => {
        const nums = S.cols.map((c, i) => c.type === 'number' ? i : -1).filter(i => i >= 0 && S.profile[i] && S.profile[i].blank > 0);
        if (!nums.length) { toast('没有需要填充的数值列', { kind: 'info' }); return; }
        const ok = await confirmDlg({ title: '批量填充', message: '将把 ' + nums.length + ' 个数值列的缺失值全部填为 0。', detail: nums.map(i => S.cols[i].name + '（' + S.profile[i].blank + ' 个）').join('、'), okText: '开始填充' });
        if (!ok) return;
        snapshot('所有数值列缺失值填 0');
        let n = 0;
        nums.forEach(ci => { for (const r of S.rows) if (isBlank(r[ci])) { r[ci] = 0; n++; } });
        afterMutation('已为 ' + nums.length + ' 个数值列填充 ' + fmtNum(n) + ' 个 0');
      };
      D.btnRenameCol.onclick = () => renameCol(+D.colOp.value);
      D.btnTypeNum.onclick = () => convertType(+D.colOp.value, 'number');
      D.btnTypeDate.onclick = () => convertType(+D.colOp.value, 'date');
      D.btnTypeText.onclick = () => convertType(+D.colOp.value, 'text');
      D.btnDropCol.onclick = () => dropCol(+D.colOp.value);
      D.btnSplit.onclick = () => splitColumn(+D.splitCol.value, D.splitSep.value || '-');
      D.btnDedupe.onclick = () => dedupeBy(+D.dedupeCol.value);
      D.btnCalcPreview.onclick = () => calcPreview();
      D.btnCalcApply.onclick = () => applyCalc();
      D.btnCalcHelp.onclick = calcHelp;
      D.calcExpr.oninput = debounce(() => { if (D.calcExpr.value.trim()) calcPreview(); else D.calcPreview.innerHTML = ''; }, 400);
    }

/* =========================================================
       报告生成与导出
       ========================================================= */
    function buildReportIfActive() { if (currentView === 'report') buildReport(); }

    function reportHTML(forExport) {
      if (!S.rows.length) return '';
      const inc = {
        kpi: D.rpIncKpi.checked, profile: D.rpIncProfile.checked, issues: D.rpIncIssues.checked,
        charts: D.rpIncCharts.checked, pivot: D.rpIncPivot.checked, sample: D.rpIncSample.checked
      };
      const title = D.rpTitle.value.trim() || ((S.file.name || '数据').replace(/\.[^.]+$/, '') + ' 分析报告');
      const author = D.rpAuthor.value.trim();
      const summary = D.rpSummary.value.trim();
      const now = new Date();
      const P = S.profile || [];
      const cells = S.rows.length * S.cols.length;
      const miss = P.reduce((s, p) => s + p.blank, 0);
      const dup = countDupRows();
      const mix = countTypes(P);
      let h = '';

      h += '<h1>' + esc(title) + '</h1>';
      h += '<div class="rp-meta">数据源：' + esc(S.file.name || '未命名') + (S.file.sheet ? '（工作表：' + esc(S.file.sheet) + '）' : '') +
        ' · ' + fmtNum(S.rows.length) + ' 行 × ' + S.cols.length + ' 列 · 生成时间：' + now.toLocaleString('zh-CN') +
        (author ? ' · ' + esc(author) : '') + '</div>';
      if (summary) h += '<p style="margin-top:12px;font-size:13px;color:#334155;line-height:1.75">' + esc(summary).replace(/\n/g, '<br>') + '</p>';

      if (inc.kpi) {
        h += '<div class="rp-kpis">' +
          '<div class="rp-kpi"><div class="l">数据行数</div><div class="v">' + fmtNum(S.rows.length) + '</div></div>' +
          '<div class="rp-kpi"><div class="l">字段数</div><div class="v">' + S.cols.length + '</div></div>' +
          '<div class="rp-kpi"><div class="l">数据完整度</div><div class="v">' + (cells ? (100 - miss / cells * 100).toFixed(1) : '100.0') + '%</div></div>' +
          '<div class="rp-kpi"><div class="l">重复行</div><div class="v">' + fmtNum(dup) + '</div></div>' +
          '</div>';
        h += '<p style="margin-top:10px;font-size:12px;color:#64748b">字段构成：' + mix.map(m => TYPE_LABEL[m.t] + ' ' + m.n + ' 列').join(' · ') + '</p>';
      }

      if (inc.profile && P.length) {
        h += '<h2>一、字段画像</h2><table class="rp-tbl"><thead><tr><th>字段</th><th>类型</th><th>缺失</th><th>唯一值</th><th>关键统计</th></tr></thead><tbody>';
        P.forEach(p => {
          let stat = '—';
          if (p.type === 'number' && p.count) stat = '最小 ' + fmtNum(p.min) + ' / 中位 ' + fmtNum(p.median) + ' / 均值 ' + fmtNum(p.mean, 2) + ' / 最大 ' + fmtNum(p.max) + ' / 合计 ' + fmtCompact(p.sum);
          else if (p.type === 'date' && p.dmin !== null && p.dmin !== undefined) stat = fmtDate(new Date(p.dmin)) + ' ~ ' + fmtDate(new Date(p.dmax)) + '（' + (p.spanDays || 0) + ' 天）';
          else if (p.top && p.top.length) stat = '高频值：' + p.top.map(t => esc(String(t[0]).slice(0, 12)) + '(' + t[1] + ')').join('、');
          h += '<tr><td>' + esc(p.name) + '</td><td>' + TYPE_LABEL[p.type] + '</td><td>' + (p.missing * 100).toFixed(1) + '%</td><td>' + fmtNum(p.unique) + '</td><td>' + stat + '</td></tr>';
        });
        h += '</tbody></table>';
      }

      if (inc.issues && S.issues.length) {
        h += '<h2>' + (inc.profile ? '二' : '一') + '、数据质量提醒</h2><ul style="margin:6px 0 0 18px;font-size:12.5px;color:#334155;line-height:1.9">';
        S.issues.slice(0, 12).forEach(i => { h += '<li><b>' + esc(i.t) + '</b>：' + esc(i.d) + '</li>'; });
        h += '</ul>';
      }

      let sec = (inc.profile ? 1 : 0) + (inc.issues ? 1 : 0);
      const CN = ['', '一', '二', '三', '四', '五', '六', '七'];
      if (inc.charts && S.saved.length) {
        sec++;
        h += '<h2>' + (CN[sec] || sec) + '、数据可视化</h2>';
        S.saved.forEach((s, i) => {
          h += '<div class="rp-chart"><img src="' + s.img + '" alt="' + esc(s.title) + '"><div class="cap">图 ' + (i + 1) + '　' + esc(s.title) + (s.meta ? '　|　' + esc(s.meta) : '') + '</div></div>';
        });
      } else if (inc.charts) {
        sec++;
        h += '<h2>' + (CN[sec] || sec) + '、数据可视化</h2><p style="font-size:12.5px;color:#94a3b8">尚未固定任何图表。请在「可视化」页面配置图表后点击「固定」。</p>';
      }

      if (inc.pivot && S.pv.last) {
        sec++;
        h += '<h2>' + (CN[sec] || sec) + '、透视分析：' + esc(S.pv.last.title) + '</h2>';
        h += '<p style="font-size:12px;color:#64748b;margin-bottom:6px">' + esc(S.pv.last.info) + '</p>';
        const m = S.pv.last;
        h += '<table class="rp-tbl"><thead><tr>' + m.head.map(x => '<th>' + esc(x) + '</th>').join('') + '</tr></thead><tbody>' +
          m.matrix.slice(0, 60).map(r => '<tr>' + r.map(v => '<td>' + esc(typeof v === 'number' ? fmtNum(v) : String(v)) + '</td>').join('') + '</tr>').join('') +
          '</tbody></table>' + (m.matrix.length > 60 ? '<p style="font-size:11.5px;color:#94a3b8;margin-top:5px">仅展示前 60 行，完整结果请导出 Excel。</p>' : '');
      }

      if (inc.sample) {
        sec++;
        const n = parseInt(D.rpSampleRows.value, 10) || 25;
        const cs = visibleCols().slice(0, 10);
        const rs = S.view.idx.slice(0, n);
        h += '<h2>' + (CN[sec] || sec) + '、数据样本（前 ' + rs.length + ' 行' + (S.view.idx.length !== S.rows.length ? '，已应用筛选' : '') + '）</h2>';
        h += '<table class="rp-tbl"><thead><tr><th>#</th>' + cs.map(i => '<th>' + esc(S.cols[i].name) + '</th>').join('') + '</tr></thead><tbody>' +
          rs.map((ri, k) => '<tr><td>' + (k + 1) + '</td>' + cs.map(i => '<td>' + esc(shortLabel(fmtVal(S.rows[ri][i], S.cols[i].type), 26)) + '</td>').join('') + '</tr>').join('') +
          '</tbody></table>';
      }

      h += '<p style="margin-top:26px;padding-top:10px;border-top:1px solid #e8ebf3;font-size:11px;color:#94a3b8">' +
        '本报告由 DataLens Pro 在浏览器本地生成 · 所有计算均在您的设备上完成，数据未上传至任何服务器</p>';
      return h;
    }

    function buildReport() {
      if (!S.rows.length) {
        const e = emptyBox('i-report', '还没有可写入报告的数据', '导入数据并在「可视化」中固定图表后，这里会生成一份完整报告。', '<button class="btn btn-primary btn-sm" onclick="go(\'ingest\')">去导入数据</button>');
        D.reportPreview.innerHTML = e; D.reportPrintArea.innerHTML = '';
        return;
      }
      const inner = reportHTML(false);
      D.reportPreview.innerHTML = '<div class="report-page">' + inner + '</div>';
      D.reportPrintArea.innerHTML = '<div class="report-page">' + inner + '</div>';
    }

    const REPORT_CSS = '.report-page{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;color:#0f172a;max-width:900px;margin:0 auto;padding:32px;background:#fff}' +
      '.report-page h1{font-size:25px;margin:0;letter-spacing:-.4px}.report-page .rp-meta{color:#64748b;font-size:12.5px;margin-top:6px}' +
      '.report-page h2{font-size:16px;margin:28px 0 10px;padding-bottom:7px;border-bottom:2px solid #eef0f6}' +
      '.rp-kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:18px}' +
      '.rp-kpi{border:1px solid #e8ebf3;border-radius:12px;padding:12px 14px;background:#fbfcfe}.rp-kpi .l{font-size:11.5px;color:#64748b}.rp-kpi .v{font-size:21px;font-weight:750;margin-top:3px}' +
      '.rp-chart{margin-top:14px;border:1px solid #e8ebf3;border-radius:12px;overflow:hidden;page-break-inside:avoid}.rp-chart img{width:100%;display:block}' +
      '.rp-chart .cap{padding:8px 12px;font-size:12px;color:#64748b;background:#fbfcfe;border-top:1px solid #eef0f6}' +
      'table.rp-tbl{width:100%;border-collapse:collapse;font-size:12px;margin-top:10px}table.rp-tbl th,table.rp-tbl td{border:1px solid #e8ebf3;padding:6px 9px;text-align:left}table.rp-tbl th{background:#f6f8fc;font-weight:650}table.rp-tbl tr:nth-child(even) td{background:#fcfdff}' +
      '@media print{body{margin:0}.report-page{padding:0;max-width:none}}';

    function exportHtmlReport() {
      if (!S.rows.length) { toast('请先导入数据', { kind: 'warn' }); return; }
      const body = reportHTML(true);
      const title = D.rpTitle.value.trim() || ((S.file.name || '数据').replace(/\.[^.]+$/, '') + ' 分析报告');
      const doc = '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">' +
        '<meta name="viewport" content="width=device-width,initial-scale=1"><title>' + esc(title) + '</title>' +
        '<style>' + REPORT_CSS + '</style></head><body><div class="report-page">' + body + '</div></body></html>';
      downloadText(doc, stampName(title.replace(/[\\/:*?"<>|]/g, '_').slice(0, 40) || '数据分析报告', 'html'), 'text/html');
      toast('报告已导出为独立 HTML 文件，可直接发送或打印', { kind: 'ok', title: '导出成功' });
    }

    function bindReportUI() {
      ['rpTitle', 'rpAuthor', 'rpSummary'].forEach(id => D[id].addEventListener('input', debounce(buildReport, 320)));
      ['rpIncKpi', 'rpIncProfile', 'rpIncIssues', 'rpIncCharts', 'rpIncPivot', 'rpIncSample'].forEach(id => D[id].addEventListener('change', () => {
        D.rpSampleWrap.style.display = D.rpIncSample.checked ? '' : 'none';
        buildReport();
      }));
      D.rpSampleRows.onchange = buildReport;
      D.btnReportRefresh.onclick = () => { if (S.rows.length) renderPivot(); buildReport(); toast('报告预览已刷新', { kind: 'ok', ms: 1600 }); };
      D.btnPrint.onclick = () => {
        if (!S.rows.length) { toast('请先导入数据', { kind: 'warn' }); return; }
        go('report'); buildReport();
        setTimeout(() => window.print(), 260);
      };
      D.btnExportHtml.onclick = exportHtmlReport;
      document.querySelectorAll('#view-report [data-exp]').forEach(b => b.onclick = () => {
        const v = b.dataset.exp;
        if (v === 'csv') exportData('csv', 'filtered');
        else if (v === 'csv-all') exportData('csv', 'all');
        else if (v === 'xlsx') exportData('xlsx', 'all');
        else if (v === 'json') exportData('json', 'all');
      });
      D.rpSampleWrap.style.display = D.rpIncSample.checked ? '' : 'none';
    }

    /* =========================================================
       本地存储：设置 / 会话 / 最近文件
       ========================================================= */
    const LS = { set: 'datalens.settings', data: 'datalens.data', recent: 'datalens.recent' };
    function saveSettings() { try { localStorage.setItem(LS.set, JSON.stringify(S.set)); } catch (e) { } }
    function loadSettings() {
      try {
        const s = JSON.parse(localStorage.getItem(LS.set) || '{}');
        Object.keys(s).forEach(k => { if (S.set.hasOwnProperty(k)) S.set[k] = s[k]; });
      } catch (e) { }
      S.view.size = S.set.pageSize || 25;
      S.view.density = S.set.density || 'normal';
      S.cfg.palette = S.set.palette || 0;
    }
    const maybeAutosave = debounce(() => {
      if (!S.set.autosave || !S.rows.length) return;
      if (S.rows.length * S.cols.length > 1200000) return;
      try {
        localStorage.setItem(LS.data, JSON.stringify({
          cols: S.cols, rows: S.rows, file: S.file, at: Date.now()
        }));
      } catch (e) {
        toast('浏览器存储空间不足，本次会话未能自动保存', { kind: 'warn', ms: 5000 });
        S.set.autosave = false; D.optAutosave.checked = false; saveSettings();
      }
    }, 1600);

    function addRecent() {
      try {
        const list = JSON.parse(localStorage.getItem(LS.recent) || '[]');
        const item = { name: S.file.name, size: S.file.size, rows: S.rows.length, cols: S.cols.length, at: Date.now() };
        const out = list.filter(x => x.name !== item.name);
        out.unshift(item);
        localStorage.setItem(LS.recent, JSON.stringify(out.slice(0, 8)));
        renderRecent();
      } catch (e) { }
    }
    function renderRecent() {
      let list = [];
      try { list = JSON.parse(localStorage.getItem(LS.recent) || '[]'); } catch (e) { }
      if (!list.length) {
        D.recentList.innerHTML = '<div class="tiny dim" style="padding:8px 4px">还没有导入记录。导入文件后会在这里留下痕迹（仅保存在本机浏览器中）。</div>';
        return;
      }
      D.recentList.innerHTML = list.map((r, i) =>
        '<div class="row gap-3" style="padding:8px 6px;border-bottom:1px dashed var(--border)">' +
        '<span class="kpi-ic" style="width:30px;height:30px;border-radius:9px"><svg class="ic ic-sm"><use href="#i-file"/></svg></span>' +
        '<div class="grow" style="min-width:0"><div class="tiny b truncate">' + esc(r.name || '未命名') + '</div>' +
        '<div class="xs dim">' + fmtNum(r.rows) + ' 行 × ' + r.cols + ' 列' + (r.size ? ' · ' + fmtBytes(r.size) : '') + ' · ' + new Date(r.at).toLocaleString('zh-CN') + '</div></div>' +
        '<button class="btn btn-xs btn-danger" data-del="' + i + '">移除</button></div>').join('');
      D.recentList.querySelectorAll('[data-del]').forEach(b => b.onclick = () => {
        try {
          const l = JSON.parse(localStorage.getItem(LS.recent) || '[]');
          l.splice(+b.dataset.del, 1); localStorage.setItem(LS.recent, JSON.stringify(l)); renderRecent();
        } catch (e) { }
      });
    }

    function restoreLast() {
      let d = null;
      try { d = JSON.parse(localStorage.getItem(LS.data) || 'null'); } catch (e) { }
      if (!d || !d.rows || !d.rows.length) { toast('没有找到可恢复的会话数据（请在导入后开启「自动保存本次会话」）', { kind: 'info', ms: 5000 }); return; }
      progress(true);
      setTimeout(() => {
        try {
          S.cols = d.cols.map(c => ({ name: c.name, type: c.type }));
          S.rows = d.rows;
          S.origin = { cols: S.cols.map(c => ({ ...c })), rows: S.rows.map(r => r.slice()) };
          S.hist = []; S.future = []; S.sel.clear(); S.saved = [];
          S.view.filters = {}; S.view.search = ''; S.view.tokens = []; S.view.sort = { c: -1, d: 0 }; S.view.page = 1; S.view.hidden = [];
          S.file = Object.assign({ name: '恢复的会话', size: 0, sheets: [], sheet: '', at: Date.now(), source: '本地缓存' }, d.file || {});
          S.cols.forEach((c, i) => { c.type = inferType(S.rows.map(r => r[i])); });
          bumpVer();
          afterLoad();
          toast('已恢复 ' + fmtNum(S.rows.length) + ' 行数据', { kind: 'ok', title: '恢复成功' });
        } catch (e) {
          console.error(e); toast('恢复失败：' + (e.message || e), { kind: 'err' });
        } finally { progress(false); }
      }, 30);
    }

/* =========================================================
       主题 / 设置 / 帮助 / 命令面板 / 快捷键
       ========================================================= */
    function darkBySystem() { return window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches; }
    function applyTheme(pref, quiet) {
      S.set.theme = pref;
      document.documentElement.setAttribute('data-theme-pref', pref);
      const dark = pref === 'dark' || (pref === 'system' && darkBySystem());
      document.documentElement.setAttribute('data-theme', dark ? 'dark' : 'light');
      const u = D.themeIcon && D.themeIcon.querySelector('use');
      if (u) u.setAttribute('href', dark ? '#i-sun' : '#i-moon');
      D.btnTheme.setAttribute('data-tip', '当前：' + (pref === 'system' ? '跟随系统' : pref === 'dark' ? '深色' : '浅色') + '（点击切换）');
      saveSettings();
      if (!quiet) {
        if (S.chart && S.chartOpt) { try { S.chart.setOption(buildChartThemePatch(), false); renderChart(); } catch (e) { renderChart(); } }
        else renderChart();
      }
    }
    function buildChartThemePatch() { return {}; }
    function cycleTheme() {
      const order = ['light', 'dark', 'system'];
      const i = order.indexOf(S.set.theme);
      const next = order[(i + 1) % order.length];
      applyTheme(next);
      toast('主题：' + (next === 'system' ? '跟随系统' : next === 'dark' ? '深色' : '浅色'), { kind: 'ok', ms: 1600 });
    }

    function showSettings() {
      const m = openModal({
        title: '偏好设置', icon: 'i-sliders',
        body:
          '<div class="col gap-4">' +
          '<div class="field"><label>外观主题</label><div class="seg" data-theme-seg style="width:100%">' +
          ['light|浅色', 'dark|深色', 'system|跟随系统'].map(x => { const a = x.split('|'); return '<button data-th="' + a[0] + '" class="grow center' + (S.set.theme === a[0] ? ' on' : '') + '">' + a[1] + '</button>'; }).join('') +
          '</div></div>' +
          '<div class="grid g-2">' +
          '<div class="field"><label for="stPageSize">表格默认每页行数</label><select class="select" id="stPageSize">' +
          [10, 25, 50, 100, 500, 0].map(v => '<option value="' + v + '"' + (S.set.pageSize === v ? ' selected' : '') + '>' + (v === 0 ? '全部（不分页）' : v) + '</option>').join('') + '</select></div>' +
          '<div class="field"><label for="stDensity">表格行高</label><select class="select" id="stDensity">' +
          [['compact', '紧凑'], ['normal', '标准'], ['cozy', '宽松']].map(v => '<option value="' + v[0] + '"' + (S.set.density === v[0] ? ' selected' : '') + '>' + v[1] + '</option>').join('') + '</select></div>' +
          '</div>' +
          '<div class="field"><label>默认图表配色</label><div class="swatches" data-pal>' +
          PALETTES.map((p, i) => '<button class="sw ' + (S.set.palette === i ? 'on' : '') + '" data-p="' + i + '" data-tip="' + p.name + '" style="background:linear-gradient(135deg,' + p.c[0] + ',' + p.c[1] + ' 50%,' + p.c[2] + ')"></button>').join('') +
          '</div></div>' +
          '<div class="field"><label>数据与隐私</label>' +
          '<label class="switch tiny" style="margin-top:4px"><input type="checkbox" id="stAuto"' + (S.set.autosave ? ' checked' : '') + '><span class="track"></span>自动保存本次会话到浏览器（便于下次恢复）</label>' +
          '<p class="hint" style="margin-top:6px">所有解析、统计、图表渲染都在你的浏览器中完成，不会向任何服务器发送数据。关闭页面后，除非开启了自动保存，否则数据即被丢弃。</p></div>' +
          '<div class="field"><label>本地缓存</label><div class="row gap-2 wrap">' +
          '<button class="btn btn-sm" id="stClearData"><svg class="ic ic-sm"><use href="#i-eraser"/></svg>清除已保存的会话数据</button>' +
          '<button class="btn btn-sm" id="stClearRecent"><svg class="ic ic-sm"><use href="#i-trash"/></svg>清除导入记录</button>' +
          '</div></div>' +
          '</div>',
        footer: '<button class="btn btn-primary" data-close>完成</button>',
        onMount(root, api) {
          root.querySelectorAll('[data-th]').forEach(b => b.onclick = () => {
            applyTheme(b.dataset.th);
            root.querySelectorAll('[data-th]').forEach(x => x.classList.toggle('on', x === b));
          });
          root.querySelector('#stPageSize').onchange = e => {
            S.set.pageSize = parseInt(e.target.value, 10) || 0; saveSettings();
            S.view.size = S.set.pageSize; D.pageSize.value = String(S.view.size); S.view.page = 1; renderTable();
          };
          root.querySelector('#stDensity').onchange = e => {
            S.set.density = e.target.value; S.view.density = e.target.value; saveSettings();
            $$('#tblDensity button').forEach(x => x.classList.toggle('on', x.dataset.d === e.target.value));
            renderTable();
          };
          root.querySelectorAll('[data-pal]').forEach(b => b.onclick = () => {
            S.set.palette = +b.dataset.p; S.cfg.palette = S.set.palette; saveSettings();
            root.querySelectorAll('[data-p]').forEach(x => x.classList.toggle('on', x === b));
            renderPalettes(); renderChart();
          });
          root.querySelector('#stAuto').onchange = e => {
            S.set.autosave = e.target.checked; D.optAutosave.checked = e.target.checked; saveSettings();
            if (e.target.checked) { maybeAutosave(); toast('已开启自动保存', { kind: 'ok', ms: 1800 }); }
            else { try { localStorage.removeItem(LS.data); } catch (err) { } toast('已关闭自动保存并清除缓存数据', { kind: 'info', ms: 2200 }); }
          };
          root.querySelector('#stClearData').onclick = () => { try { localStorage.removeItem(LS.data); } catch (e) { } toast('已清除本地会话数据', { kind: 'ok' }); };
          root.querySelector('#stClearRecent').onclick = () => { try { localStorage.removeItem(LS.recent); } catch (e) { } renderRecent(); toast('已清除导入记录', { kind: 'ok' }); };
        }
      });
      return m;
    }

    const SHORTCUTS = [
      ['Ctrl / ⌘ + K', '打开命令面板', 'i-command'], ['Ctrl / ⌘ + O', '导入数据文件', 'i-upload'],
      ['Ctrl / ⌘ + S', '导出当前筛选结果为 CSV', 'i-download'], ['Ctrl / ⌘ + F', '定位到数据表并搜索', 'i-search'],
      ['Ctrl / ⌘ + Z', '撤销上一步清洗操作', 'i-undo'], ['Ctrl / ⌘ + ⇧ + Z', '重做', 'i-refresh'],
      ['/', '快速聚焦搜索框', 'i-search'], ['?', '打开本帮助', 'i-help'],
      ['Esc', '关闭弹层 / 退出全屏', 'i-x'], ['双击单元格', '直接编辑数据', 'i-pencil'],
      ['右键单元格', '复制 / 筛选 / 排序菜单', 'i-dots'], ['拖拽文件到窗口', '任意位置即可导入', 'i-upload']
    ];
    function showHelp() {
      openModal({
        title: '帮助与快捷键', icon: 'i-help', size: 'wide',
        body:
          '<div class="grid g-2" style="gap:18px">' +
          '<div><h4 class="b" style="font-size:14px;margin-bottom:10px">键盘与鼠标</h4><div class="kbdlist">' +
          SHORTCUTS.map(s => '<div class="kbdrow"><svg class="ic ic-sm" style="color:var(--text-3)"><use href="#' + s[2] + '"/></svg>' + esc(s[1]) + '<span class="keys"><span class="kbd">' + esc(s[0]) + '</span></span></div>').join('') +
          '</div></div>' +
          '<div><h4 class="b" style="font-size:14px;margin-bottom:10px">典型工作流</h4>' +
          [['1', '导入', '把 CSV / Excel / JSON 拖进窗口；多工作表文件可在数据源卡片中切换。'],
          ['2', '体检', '在「数据概览」查看每列类型、缺失率、分布与质量问题提醒。'],
          ['3', '清洗', '在「数据清洗」去重、填充缺失、转换类型或新增计算列，每步都可撤销。'],
          ['4', '分析', '在「数据表」中筛选与排序，在「可视化」中做聚合图表，在「透视分析」中交叉汇总。'],
          ['5', '交付', '固定满意的图表，到「报告导出」一键生成可打印 / 可分享的 HTML 报告。']].map(t =>
            '<div class="tip-step"><span class="tip-n">' + t[0] + '</span><div><div class="tiny b">' + esc(t[1]) + '</div><div class="tiny muted" style="margin-top:2px">' + esc(t[2]) + '</div></div></div>').join('') +
          '<h4 class="b" style="font-size:14px;margin:16px 0 8px">关于</h4>' +
          '<p class="tiny muted" style="line-height:1.8">DataLens Pro 是一个纯前端单文件应用：HTML + CSS + JavaScript，无后端、无构建、无追踪。' +
          '表格解析使用 SheetJS，图表使用 Apache ECharts（均通过 CDN 加载，首次使用需联网）。' +
          '所有数据处理都在你的浏览器内完成。</p>' +
          '<p class="tiny dim" style="margin-top:8px">版本 2.1 · 支持 CSV / TSV / XLSX / XLS / JSON · 最大建议 50 MB</p>' +
          '</div></div>',
        footer: '<button class="btn" data-close>关闭</button><button class="btn btn-primary" data-demo>载入示例数据试试</button>',
        onMount(root, api) { root.querySelector('[data-demo]').onclick = () => { api.close(); loadDemo(); }; }
      });
    }

    function showIntro() {
      S.set.seenIntro = true; saveSettings();
      openModal({
        title: '欢迎使用 DataLens Pro', icon: 'i-sparkles',
        body: '<p class="muted" style="font-size:13.5px">这是一个完全在本机运行的数据分析工作台。三步开始：</p>' +
          '<div style="margin-top:12px">' +
          [['1', '把数据拖进来', '支持 CSV、Excel（多工作表）、TSV 与 JSON，首行自动识别为列名。'],
          ['2', '让它自己体检', '自动生成列画像：类型、缺失率、唯一值、极值、分布与质量问题清单。'],
          ['3', '一键出图出报告', '智能推荐图表配置，固定图表后可导出可打印的 HTML 报告。']].map(t =>
            '<div class="tip-step"><span class="tip-n">' + t[0] + '</span><div><div class="tiny b">' + esc(t[1]) + '</div><div class="tiny muted" style="margin-top:2px">' + esc(t[2]) + '</div></div></div>').join('') +
          '</div>' +
          '<div class="libwarn" style="margin-top:10px"><svg class="ic ic-sm" style="color:var(--warn);margin-top:2px"><use href="#i-lock"/></svg><div>隐私说明：文件不会被上传到任何服务器，全部计算在浏览器中完成。关闭页面即清除数据（除非你开启自动保存）。</div></div>',
        footer: '<button class="btn" data-close>先自己看看</button>' +
          '<button class="btn" data-open><svg class="ic ic-sm"><use href="#i-upload"/></svg>导入文件</button>' +
          '<button class="btn btn-primary" data-demo><svg class="ic ic-sm"><use href="#i-zap"/></svg>载入示例数据</button>',
        onMount(root, api) {
          root.querySelector('[data-demo]').onclick = () => { api.close(); loadDemo(); };
          root.querySelector('[data-open]').onclick = () => { api.close(); D.fileInput.click(); };
        }
      });
    }

    /* ---------- 命令面板 ---------- */
    const VIEW_META = {
      ingest: ['数据接入', 'i-upload'], overview: ['数据概览', 'i-gauge'], table: ['数据表', 'i-table'],
      charts: ['可视化', 'i-chart'], pivot: ['透视分析', 'i-grid'], clean: ['数据清洗', 'i-sparkles'], report: ['报告导出', 'i-report']
    };
    function paletteActions() {
      const a = [], has = S.rows.length > 0;
      const add = (g, name, ic, fn, kw) => a.push({ g, name, ic, fn, kw: kw || '' });
      add('数据', '导入文件…', 'i-upload', () => D.fileInput.click(), 'open file csv excel 导入');
      add('数据', '粘贴表格 / JSON…', 'i-paste', openPasteDlg, 'paste 粘贴');
      add('数据', '载入示例数据', 'i-zap', loadDemo, 'demo 示例 演示');
      if (has) {
        add('数据', '恢复上次会话数据', 'i-history', restoreLast, 'restore 恢复');
        add('数据', '导出 CSV（当前筛选）', 'i-download', () => exportData('csv', 'filtered'), 'export csv 导出');
        add('数据', '导出 Excel（全部）', 'i-table', () => exportData('xlsx', 'all'), 'export xlsx excel 导出');
        add('数据', '导出 JSON', 'i-braces', () => exportData('json', 'all'), 'export json 导出');
        add('数据', '切换工作表…', 'i-layers', showSheetPicker, 'sheet 工作表');
        add('数据', '清空当前数据', 'i-trash', clearDataConfirm, 'clear 清空');
      }
      Object.keys(VIEW_META).forEach(v => add('跳转', '前往：' + VIEW_META[v][0], VIEW_META[v][1], () => go(v), 'go view 跳转 ' + v));
      if (has) {
        add('分析', '智能推荐图表配置', 'i-wand', () => { go('charts'); autoConfigChart(); }, 'auto chart 推荐');
        add('分析', '固定当前图表', 'i-pin', () => { go('charts'); saveChart(); }, 'save chart 固定');
        add('分析', '全屏显示图表', 'i-expand', () => { go('charts'); setTimeout(() => { if (!S.full) toggleFull(); }, 150); }, 'fullscreen 全屏');
        add('分析', '生成透视表', 'i-grid', () => { go('pivot'); renderPivot(); }, 'pivot 透视');
        add('分析', '清除全部筛选', 'i-eraser', () => { go('table'); clearAllFilters(); }, 'filter 筛选 清除');
        add('分析', '删除重复行', 'i-copy', () => runClean('dropDupRows'), 'dedupe 去重');
        add('报告', '生成并预览报告', 'i-report', () => { go('report'); buildReport(); }, 'report 报告');
        add('报告', '导出 HTML 报告', 'i-save', exportHtmlReport, 'export html 报告');
        add('报告', '打印 / 存为 PDF', 'i-print', () => { go('report'); buildReport(); setTimeout(() => window.print(), 260); }, 'print pdf 打印');
      }
      add('视图', '切换主题（浅色 / 深色 / 跟随系统）', 'i-moon', cycleTheme, 'theme dark light 主题');
      add('设置', '偏好设置…', 'i-sliders', showSettings, 'settings 设置 偏好');
      add('帮助', '快捷键与帮助', 'i-keyboard', showHelp, 'help shortcut 帮助 快捷键');
      add('帮助', '计算列语法说明', 'i-fx', () => { go('clean'); calcHelp(); }, 'calc formula 计算列 公式');
      add('帮助', '关于 DataLens Pro', 'i-info', showAbout, 'about 关于');
      return a;
    }
    function showAbout() {
      const st = window.__libStatus || {};
      openModal({
        title: '关于 DataLens Pro', icon: 'i-info', size: 'slim',
        body: '<div class="col gap-3">' +
          '<div class="row gap-3"><div class="brand-logo"><svg class="ic ic-lg"><use href="#i-gauge"/></svg></div><div><div class="b" style="font-size:15px">DataLens Pro <span class="badge badge-brand">v2.1</span></div>' +
          '<div class="tiny dim">纯前端数据分析与报表工作台</div></div></div>' +
          '<div class="stat-line"><span class="muted">表格解析 SheetJS</span><b>' + (st.XLSX ? '已加载' : (window.XLSX ? '已加载' : '未加载')) + '</b></div>' +
          '<div class="stat-line"><span class="muted">图表引擎 ECharts</span><b>' + ((st.echarts || window.echarts) ? (window.echarts ? window.echarts.version || '5.x' : '已加载') : '未加载') + '</b></div>' +
          '<div class="stat-line"><span class="muted">当前主题</span><b>' + (S.set.theme === 'system' ? '跟随系统' : S.set.theme === 'dark' ? '深色' : '浅色') + '</b></div>' +
          '<div class="stat-line"><span class="muted">当前数据</span><b>' + (S.rows.length ? fmtNum(S.rows.length) + ' × ' + S.cols.length : '未载入') + '</b></div>' +
          '<div class="stat-line"><span class="muted">本地缓存</span><b>' + (localStorage.getItem(LS.data) ? '有会话数据' : '无') + '</b></div>' +
          '<p class="hint" style="margin-top:6px">数据全程留在本机浏览器，不会上传到任何服务器。</p></div>',
        footer: '<button class="btn btn-primary" data-close>好的</button>'
      });
    }

    let cmdkItems = [], cmdkCur = 0;
    function openPalette() {
      if (D.cmdkRoot.firstChild) return;
      cmdkItems = paletteActions(); cmdkCur = 0;
      D.cmdkRoot.innerHTML = '<div class="mask" style="align-items:flex-start;padding-top:12vh"><div class="cmdk">' +
        '<div class="cmdk-in"><svg class="ic" style="color:var(--text-3)"><use href="#i-search"/></svg>' +
        '<input id="cmdkInput" placeholder="输入命令或搜索功能…（例如：导出、透视、深色）" autocomplete="off" data-autofocus>' +
        '<span class="kbd">Esc</span></div>' +
        '<div class="cmdk-list" id="cmdkList"></div>' +
        '<div class="cmdk-foot"><span><span class="kbd">↑</span><span class="kbd">↓</span> 选择</span><span><span class="kbd">Enter</span> 执行</span><span class="grow"></span><span>' + cmdkItems.length + ' 个命令</span></div>' +
        '</div></div>';
      const inp = D.cmdkRoot.querySelector('#cmdkInput'), list = D.cmdkRoot.querySelector('#cmdkList');
      const mask = D.cmdkRoot.querySelector('.mask');
      const close = () => { D.cmdkRoot.innerHTML = ''; };
      mask.addEventListener('mousedown', e => { if (e.target === mask) close(); });
      function score(a, q) {
        const name = a.name.toLowerCase(), kw = (a.kw || '').toLowerCase(), g = a.g.toLowerCase();
        if (name.indexOf(q) === 0) return 100;
        if (name.indexOf(q) >= 0) return 80;
        if (kw.indexOf(q) >= 0) return 60;
        if (g.indexOf(q) >= 0) return 40;
        let i = 0; for (const ch of q) { const p = name.indexOf(ch, i); if (p < 0) return 0; i = p + 1; }
        return 20;
      }
      function draw() {
        const q = inp.value.trim().toLowerCase();
        const items = q ? cmdkItems.map(a => ({ a, s: score(a, q) })).filter(x => x.s > 0).sort((x, y) => y.s - x.s).map(x => x.a) : cmdkItems;
        cmdkCur = clamp(cmdkCur, 0, Math.max(0, items.length - 1));
        let lastG = '';
        list.innerHTML = items.map((a, i) => {
          const gh = a.g !== lastG ? '<div class="pop-head">' + esc(a.g) + '</div>' : '';
          lastG = a.g;
          return gh + '<div class="cmdk-item' + (i === cmdkCur ? ' cur' : '') + '" data-i="' + i + '">' +
            '<span class="ci-ic"><svg class="ic ic-sm"><use href="#' + a.ic + '"/></svg></span>' +
            '<span class="grow truncate">' + esc(a.name) + '</span><span class="ci-group">' + esc(a.g) + '</span></div>';
        }).join('') || '<div class="empty" style="padding:26px"><p class="tiny">没有匹配的命令</p></div>';
        list._items = items;
        const cur = list.querySelector('.cmdk-item.cur');
        if (cur) cur.scrollIntoView({ block: 'nearest' });
        list.querySelectorAll('.cmdk-item').forEach(el2 => {
          el2.onclick = () => run(+el2.dataset.i);
          el2.onmouseenter = () => { cmdkCur = +el2.dataset.i; draw(); };
        });
      }
      function run(i) {
        const items = list._items || cmdkItems;
        const a = items[i]; if (!a) return;
        close(); setTimeout(() => { try { a.fn(); } catch (e) { console.error(e); toast(String(e.message || e), { kind: 'err' }); } }, 30);
      }
      inp.oninput = () => { cmdkCur = 0; draw(); };
      inp.onkeydown = e => {
        const items = list._items || cmdkItems;
        if (e.key === 'ArrowDown') { e.preventDefault(); cmdkCur = (cmdkCur + 1) % Math.max(1, items.length); draw(); }
        else if (e.key === 'ArrowUp') { e.preventDefault(); cmdkCur = (cmdkCur - 1 + items.length) % Math.max(1, items.length); draw(); }
        else if (e.key === 'Enter') { e.preventDefault(); run(cmdkCur); }
        else if (e.key === 'Escape') { e.preventDefault(); close(); }
      };
      draw();
      setTimeout(() => inp.focus(), 40);
    }

    function openPasteDlg() {
      openModal({
        title: '粘贴表格数据', icon: 'i-paste', size: 'wide',
        body: '<p class="tiny muted" style="margin-bottom:10px">从 Excel、Numbers、网页表格或文本文件直接复制，粘贴到下面。支持制表符 / 逗号 / 分号分隔，也支持 JSON 数组。</p>' +
          '<textarea class="input mono" id="pasteArea" rows="11" placeholder="姓名&#9;部门&#9;销售额&#10;张三&#9;华东&#9;12000&#10;李四&#9;华北&#9;9800" style="resize:vertical" data-autofocus></textarea>' +
          '<div class="row gap-3 wrap" style="margin-top:10px">' +
          '<div class="field" style="width:auto"><label for="pasteDelim">分隔符</label><select class="select" id="pasteDelim" style="width:auto">' +
          '<option value="auto" selected>自动识别</option><option value="\t">制表符 Tab</option><option value=",">逗号 ,</option><option value=";">分号 ;</option><option value="|">竖线 |</option></select></div>' +
          '<label class="switch tiny" style="margin-top:20px"><input type="checkbox" id="pasteHeader" checked><span class="track"></span>首行是列名</label>' +
          '<span class="grow"></span><span class="tiny dim" id="pasteInfo"></span>' +
          '</div>',
        footer: '<button class="btn" data-close>取消</button><button class="btn btn-primary" id="pasteGo">解析并导入</button>',
        onMount(root, api) {
          const ta = root.querySelector('#pasteArea'), info = root.querySelector('#pasteInfo');
          ta.oninput = () => {
            const lines = ta.value.split('\n').filter(l => l.trim());
            info.textContent = lines.length ? '检测到 ' + lines.length + ' 行 · ' + ta.value.length + ' 字符' : '';
          };
          root.querySelector('#pasteGo').onclick = () => {
            const text = ta.value;
            if (!text.trim()) { toast('请先粘贴数据', { kind: 'warn' }); return; }
            const noHeader = !root.querySelector('#pasteHeader').checked;
            progress(true);
            setTimeout(() => {
              try {
                const t = text.trim();
                if (t[0] === '[' || t[0] === '{') {
                  loadJsonText(t, { name: '粘贴的 JSON', size: text.length, source: '粘贴' });
                } else {
                  let d = root.querySelector('#pasteDelim').value;
                  if (d === 'auto') d = guessDelim(text);
                  else if (d === '\\t') d = '\t';
                  const matrix = parseDelimited(text, d);
                  if (!ingestMatrix(matrix, { name: '粘贴的数据', size: text.length, source: '粘贴', forceNoHeader: noHeader })) { progress(false); return; }
                }
                afterLoad(); api.close();
                toast('已导入粘贴的数据', { kind: 'ok', title: '导入成功' });
              } catch (e) {
                console.error(e);
                toast(String(e.message || e), { kind: 'err', title: '解析失败', ms: 7000 });
              } finally { progress(false); }
            }, 30);
          };
        }
      });
    }

/* =========================================================
       全局绑定与初始化
       ========================================================= */
    async function clearDataConfirm() {
      if (!S.rows.length) { toast('当前没有已载入的数据', { kind: 'info', ms: 1800 }); return; }
      const ok = await confirmDlg({
        title: '清空当前数据', message: '将清除已载入的 ' + fmtNum(S.rows.length) + ' 行数据、全部图表与透视结果。',
        detail: '此操作不可撤销，你可以随时重新导入文件。', okText: '确认清空', danger: true
      });
      if (ok) clearAll();
    }

    function hasFiles(e) { return e.dataTransfer && Array.prototype.indexOf.call(e.dataTransfer.types || [], 'Files') >= 0; }

    function bindGlobal() {
      $$('.nav-item').forEach(b => b.onclick = () => { if (!b.disabled) go(b.dataset.view); });
      D.navToggle.onclick = () => document.body.classList.toggle('nav-open');
      document.addEventListener('click', e => {
        if (document.body.classList.contains('nav-open') && !D.sidebar.contains(e.target) && e.target !== D.navToggle && !D.navToggle.contains(e.target)) {
          document.body.classList.remove('nav-open');
        }
      });
      D.btnTheme.onclick = cycleTheme;
      D.langToggle.onclick = toggleLanguage;
      D.btnSettings.onclick = showSettings;
      D.btnHelp.onclick = showHelp;
      D.btnPalette.onclick = openPalette;
      D.btnOpen.onclick = () => D.fileInput.click();
      D.btnDemo.onclick = loadDemo;
      D.btnDemo2.onclick = loadDemo;
      D.btnRestore.onclick = restoreLast;
      D.btnClearRecent.onclick = async () => {
        if (await confirmDlg({ title: '清除导入记录', message: '将删除本机保存的文件导入记录（不影响任何实际文件）。', okText: '清除', danger: true })) {
          try { localStorage.removeItem(LS.recent); } catch (e) { }
          renderRecent(); toast('已清除导入记录', { kind: 'ok' });
        }
      };
      D.linkHelp.onclick = e => { e.preventDefault(); showHelp(); };
      D.linkAbout.onclick = e => { e.preventDefault(); showAbout(); };
      D.optAutosave.onchange = () => {
        S.set.autosave = D.optAutosave.checked; saveSettings();
        if (S.set.autosave) { maybeAutosave(); toast('已开启：数据会保存在本机浏览器，下次可一键恢复', { kind: 'ok' }); }
        else { try { localStorage.removeItem(LS.data); } catch (e) { } toast('已关闭自动保存并清除缓存', { kind: 'info' }); }
      };
      D.btnOvToChart.onclick = () => { go('charts'); };
      D.btnReprofile.onclick = () => {
        progress(true);
        setTimeout(() => {
          const ms = reprofile();
          renderOverview(); updateSourceCard(); buildReportIfActive();
          progress(false);
          toast('已重新分析 ' + S.cols.length + ' 列，用时 ' + ms.toFixed(0) + ' ms', { kind: 'ok' });
        }, 30);
      };
      D.profSort.onclick = e => {
        const b = e.target.closest('[data-ps]'); if (!b) return;
        profSortMode = b.dataset.ps;
        $$('#profSort button').forEach(x => x.classList.toggle('on', x === b));
        renderProfile(profSortMode);
      };
      D.profBody.addEventListener('click', e => {
        const row = e.target.closest('[data-col]');
        if (row) openColDrawer(+row.dataset.col);
      });
      D.profBody.addEventListener('keydown', e => {
        const row = e.target.closest('[data-col]');
        if (row && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); openColDrawer(+row.dataset.col); }
      });

      // 拖拽导入
      D.dropzone.onclick = () => D.fileInput.click();
      D.dropzone.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); D.fileInput.click(); } };
      D.fileInput.onchange = e => { handleFiles(e.target.files); e.target.value = ''; };
      let dragDepth = 0;
      document.addEventListener('dragenter', e => { if (!hasFiles(e)) return; e.preventDefault(); dragDepth++; document.body.classList.add('dragging'); });
      document.addEventListener('dragover', e => { if (!hasFiles(e)) return; e.preventDefault(); if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy'; });
      document.addEventListener('dragleave', e => { if (!hasFiles(e)) return; dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) document.body.classList.remove('dragging'); });
      document.addEventListener('drop', e => {
        if (!hasFiles(e)) return;
        e.preventDefault(); dragDepth = 0; document.body.classList.remove('dragging');
        handleFiles(e.dataTransfer.files);
      });
      D.btnPaste.onclick = openPasteDlg;

      // 快捷键
      document.addEventListener('keydown', e => {
        const t = e.target;
        const tag = (t.tagName || '').toLowerCase();
        const typing = tag === 'input' || tag === 'textarea' || tag === 'select' || t.isContentEditable;
        const mod = e.ctrlKey || e.metaKey;
        const k = (e.key || '').toLowerCase();
        if (mod && k === 'k') { e.preventDefault(); openPalette(); return; }
        if (mod && k === 'o') { e.preventDefault(); D.fileInput.click(); return; }
        if (mod && k === 's') { e.preventDefault(); if (S.rows.length) exportData('csv', 'filtered'); else toast('还没有数据可导出', { kind: 'warn' }); return; }
        if (mod && k === 'f') {
          e.preventDefault();
          if (!S.rows.length) { toast('请先导入数据', { kind: 'info', ms: 1800 }); return; }
          go('table'); setTimeout(() => D.tblSearch.focus(), 140); return;
        }
        if (mod && (k === 'z' || k === 'y')) {
          if (typing) return;
          if (!S.rows.length) return;
          e.preventDefault();
          if (k === 'y' || e.shiftKey) redo(); else undo();
          return;
        }
        if (e.key === 'Escape') {
          if (S.full) { toggleFull(); return; }
          if (D.cmdkRoot.firstChild) { D.cmdkRoot.innerHTML = ''; return; }
          const m = topModal(); if (m) { m.close(); return; }
          if (D.drawerRoot.firstChild) { D.drawerRoot.innerHTML = ''; return; }
          if (curPop) { closePop(); return; }
          if (document.body.classList.contains('nav-open')) { document.body.classList.remove('nav-open'); return; }
          return;
        }
        if (typing || mod || e.altKey) return;
        if (e.key === '/') {
          if (!S.rows.length) return;
          e.preventDefault(); go('table'); setTimeout(() => D.tblSearch.focus(), 140);
        } else if (e.key === '?') { e.preventDefault(); showHelp(); }
        else if (k === 'd' && e.shiftKey) { e.preventDefault(); cycleTheme(); }
      });

      if (window.matchMedia) {
        const mq = matchMedia('(prefers-color-scheme: dark)');
        const onChange = () => { if (S.set.theme === 'system') applyTheme('system'); };
        if (mq.addEventListener) mq.addEventListener('change', onChange);
        else if (mq.addListener) mq.addListener(onChange);
      }
      window.addEventListener('resize', throttle(() => { if (S.chart && !S.chart.isDisposed()) S.chart.resize(); }, 220));
      window.addEventListener('error', ev => {
        if (ev && ev.message && /echarts|XLSX/i.test(ev.message)) console.warn('[DataLens] 库相关错误：', ev.message);
      });
    }

    function showLibWarn(st) {
      const miss = [];
      if (!st.XLSX && !window.XLSX) miss.push('表格解析库 SheetJS');
      if (!st.echarts && !window.echarts) miss.push('图表库 ECharts');
      if (!miss.length) return;
      const host = $('#view-ingest .card-b');
      if (!host || host.querySelector('#libWarnBox')) return;
      const div = document.createElement('div');
      div.className = 'libwarn'; div.id = 'libWarnBox'; div.style.marginBottom = '16px';
      div.innerHTML = '<svg class="ic" style="color:var(--warn);margin-top:2px;flex:none"><use href="#i-alert"/></svg>' +
        '<div><b>' + miss.join(' 与 ') + ' 未能加载。</b><br>可能处于离线环境或 CDN 被拦截。此时「粘贴文本数据」、表格浏览、筛选与清洗仍可使用，' +
        '但 Excel 文件解析与图表绘制需要联网。请连接网络后刷新页面重试。</div>';
      host.insertBefore(div, host.firstChild);
    }

    async function init() {
      cacheDom();
      loadSettings();
      applyLanguage(S.set.lang || 'zh', true);
      applyTheme(S.set.theme || 'system', true);
      D.optAutosave.checked = !!S.set.autosave;
      D.pageSize.value = String(S.view.size);
      $$('#tblDensity button').forEach(b => b.classList.toggle('on', b.dataset.d === S.view.density));
      bindGlobal(); bindTable(); bindChartUI(); bindPivotUI(); bindCleanUI(); bindReportUI();
      renderChartTypes(); renderPalettes(); renderSaved(); renderRecent(); renderHist(); renderCleanUI();
      enableNav(false);
      renderPivot();
      const mq = window.matchMedia ? matchMedia('(max-width: 900px)') : null;
      const syncNavBtn = () => { D.navToggle.style.display = (mq && mq.matches) ? '' : 'none'; };
      syncNavBtn();
      if (mq) { if (mq.addEventListener) mq.addEventListener('change', syncNavBtn); else if (mq.addListener) mq.addListener(syncNavBtn); }
      window.addEventListener('resize', throttle(syncNavBtn, 300));

      const st = await window.__libsReady;
      showLibWarn(st || {});
      if (st && st.XLSX && st.echarts) {
        // 库就绪后再渲染一次，确保图表可用
        if (S.rows.length) renderChart();
      }
      let hasCache = false;
      try { hasCache = !!localStorage.getItem(LS.data); } catch (e) { }
      if (hasCache) {
        toast('检测到上次保存的会话数据', { kind: 'info', title: '可以恢复', action: '立即恢复', onAction: restoreLast, ms: 9000 });
      } else if (!S.set.seenIntro) {
        setTimeout(showIntro, 420);
      } else {
        D.btnRestore.disabled = true;
        D.btnRestore.setAttribute('data-tip', '开启「自动保存本次会话」后，这里可以恢复上次的数据');
      }
      console.log('%cDataLens Pro 已就绪', 'color:#4f46e5;font-weight:700;font-size:13px', '· 按 Ctrl/⌘+K 打开命令面板');
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
