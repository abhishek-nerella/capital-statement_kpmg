/* Shared UI primitives */

const Icon = ({ name, size = 14, color = "currentColor", stroke = 1.8 }) => {
  const s = { width: size, height: size, color, flexShrink: 0 };
  const sp = { fill: "none", stroke: "currentColor", strokeWidth: stroke, strokeLinecap: "round", strokeLinejoin: "round" };
  switch (name) {
    case "upload":      return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M12 4v12M6 10l6-6 6 6M4 20h16"/></svg>);
    case "file":        return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6"/></svg>);
    case "file-x":      return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6"/><path d="M9 13l6 6M15 13l-6 6"/></svg>);
    case "check":       return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M5 12l5 5L20 7"/></svg>);
    case "x":           return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M6 6l12 12M18 6L6 18"/></svg>);
    case "warn":        return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M12 3l10 18H2L12 3z"/><path d="M12 10v4M12 18v.01"/></svg>);
    case "chevron-down":return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M6 9l6 6 6-6"/></svg>);
    case "chevron-right":return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M9 6l6 6-6 6"/></svg>);
    case "download":    return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M12 4v12M6 14l6 6 6-6M4 20h16"/></svg>);
    case "play":        return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M7 4v16l13-8L7 4z"/></svg>);
    case "spark":       return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M12 3v4M12 17v4M3 12h4M17 12h4M5.6 5.6l2.8 2.8M15.6 15.6l2.8 2.8M5.6 18.4l2.8-2.8M15.6 8.4l2.8-2.8"/></svg>);
    case "send":        return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M22 2L11 13M22 2l-7 20-4-9-9-4 20-7z"/></svg>);
    case "msg":         return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M21 12a8 8 0 0 1-12.6 6.5L3 20l1.5-5.4A8 8 0 1 1 21 12z"/></svg>);
    case "search":      return (<svg viewBox="0 0 24 24" style={s} {...sp}><circle cx="11" cy="11" r="7"/><path d="M21 21l-4-4"/></svg>);
    case "grid":        return (<svg viewBox="0 0 24 24" style={s} {...sp}><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/></svg>);
    case "list":        return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/></svg>);
    case "filter":      return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M3 4h18l-7 9v7l-4-2v-5L3 4z"/></svg>);
    case "refresh":     return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.5 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.65 4.36A9 9 0 0 0 20.5 15"/></svg>);
    case "audit":       return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13h8M8 17h5M8 9h2"/></svg>);
    case "shield":      return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M12 2l9 4v6c0 5-3.5 9-9 10-5.5-1-9-5-9-10V6l9-4z"/></svg>);
    case "bolt":        return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M13 2L4 14h6l-1 8 9-12h-6l1-8z"/></svg>);
    case "user":        return (<svg viewBox="0 0 24 24" style={s} {...sp}><circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/></svg>);
    case "users":       return (<svg viewBox="0 0 24 24" style={s} {...sp}><circle cx="9" cy="8" r="4"/><path d="M1 21a8 8 0 0 1 16 0"/><path d="M17 4a4 4 0 0 1 0 8M23 21a8 8 0 0 0-6-7.7"/></svg>);
    case "doc":         return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M9 13h6M9 17h6M9 9h1"/></svg>);
    case "calendar":    return (<svg viewBox="0 0 24 24" style={s} {...sp}><rect x="3" y="5" width="18" height="16" rx="1"/><path d="M16 3v4M8 3v4M3 11h18"/></svg>);
    case "settings":    return (<svg viewBox="0 0 24 24" style={s} {...sp}><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 0 1-4 0v-.1a1.7 1.7 0 0 0-1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 0 1 0-4h.1a1.7 1.7 0 0 0 1.5-1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3h0a1.7 1.7 0 0 0 1-1.5V3a2 2 0 0 1 4 0v.1a1.7 1.7 0 0 0 1 1.5h0a1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8v0a1.7 1.7 0 0 0 1.5 1H21a2 2 0 0 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg>);
    case "chart":       return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M3 3v18h18"/><path d="M7 14l4-4 4 4 5-6"/></svg>);
    case "table":       return (<svg viewBox="0 0 24 24" style={s} {...sp}><rect x="3" y="4" width="18" height="16" rx="1"/><path d="M3 10h18M3 16h18M10 4v16M16 4v16"/></svg>);
    case "info":        return (<svg viewBox="0 0 24 24" style={s} {...sp}><circle cx="12" cy="12" r="9"/><path d="M12 8v.01M12 12v4"/></svg>);
    case "external":    return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M14 4h6v6M20 4l-9 9M14 14v6H4V8h6"/></svg>);
    case "plus":        return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M12 5v14M5 12h14"/></svg>);
    case "trash":       return (<svg viewBox="0 0 24 24" style={s} {...sp}><path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M6 6l1 14a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-14"/></svg>);
    case "copy":        return (<svg viewBox="0 0 24 24" style={s} {...sp}><rect x="9" y="9" width="13" height="13" rx="1"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>);
    case "dot":         return (<svg viewBox="0 0 24 24" style={s} fill="currentColor"><circle cx="12" cy="12" r="4"/></svg>);
    default:            return null;
  }
};

const Pill = ({ children, color = "default", onClick, removable, onRemove }) => {
  const cls = "pill" + (color === "blue" ? " pill-blue" : "");
  return (
    <span className={cls} onClick={onClick} style={{ cursor: onClick ? "pointer" : "default" }}>
      {children}
      {removable && (
        <button onClick={(e) => { e.stopPropagation(); onRemove && onRemove(); }}
          style={{ background: "none", border: "none", color: "inherit", padding: 0, marginLeft: 2, cursor: "pointer", display: "inline-flex" }}>
          <Icon name="x" size={11} />
        </button>
      )}
    </span>
  );
};

const Badge = ({ verdict, size = "md" }) => {
  const map = {
    ALL_PASS:   { cls: "badge-pass",    label: "ALL PASS" },
    REVALUABLE: { cls: "badge-reval",   label: "REVALUABLE" },
    INVALID:    { cls: "badge-invalid", label: "INVALID" },
  };
  const m = map[verdict] || { cls: "badge-neutral", label: verdict };
  return <span className={`badge ${m.cls}`} style={{ fontSize: size === "sm" ? 9 : 10.5 }}>{m.label}</span>;
};

const MetricCard = ({ label, value, sub, accent = "light", icon }) => {
  const accentCls = accent === "pacific" ? "accent-pacific" : accent === "cobalt" ? "accent-cobalt" : accent === "teal" ? "accent-teal" : "";
  return (
    <div className={`metric ${accentCls}`}>
      <div className="m-label">{label}</div>
      <div className="m-value">{value}</div>
      {sub && <div className="m-sub">{sub}</div>}
    </div>
  );
};

const SecHeading = ({ children, right }) => (
  <div className="sec-heading" style={right ? { gap: 8 } : {}}>
    <span>{children}</span>
    <div style={{ flex: 1, height: 1, background: "var(--border)" }}></div>
    {right && <div>{right}</div>}
  </div>
);

const Collapsible = ({ title, defaultOpen = false, children, sub }) => {
  const [open, setOpen] = React.useState(defaultOpen);
  return (
    <div className="card" style={{ marginBottom: 14 }}>
      <button
        onClick={() => setOpen(!open)}
        style={{
          width: "100%", display: "flex", alignItems: "center", gap: 10,
          padding: "11px 14px", background: "var(--white)", border: "none",
          cursor: "pointer", textAlign: "left", borderBottom: open ? "1px solid var(--border)" : "none"
        }}>
        <Icon name={open ? "chevron-down" : "chevron-right"} size={14} color="var(--ink-500)" />
        <span style={{ fontWeight: 600, fontSize: 12.5, color: "var(--ink-900)" }}>{title}</span>
        {sub && <span style={{ fontSize: 11, color: "var(--ink-500)", marginLeft: 4 }}>{sub}</span>}
      </button>
      {open && <div style={{ padding: 14 }}>{children}</div>}
    </div>
  );
};

/* Validation summary tile */
const ValidationTile = ({ verdict, count, total }) => {
  const colors = {
    ALL_PASS:   { bg: "var(--teal)",    fg: "#fff" },
    REVALUABLE: { bg: "var(--pacific)", fg: "var(--dark-navy)" },
    INVALID:    { bg: "var(--purple)",  fg: "#fff" },
  }[verdict];
  return (
    <div style={{
      background: colors.bg, color: colors.fg,
      padding: "10px 12px",
      display: "flex", flexDirection: "column", gap: 4,
      minWidth: 0
    }}>
      <div style={{ fontSize: 9.5, fontWeight: 800, letterSpacing: "0.12em", opacity: 0.85 }}>
        {verdict.replace("_", " ")}
      </div>
      <div style={{ fontSize: 22, fontWeight: 800, fontVariantNumeric: "tabular-nums", lineHeight: 1 }}>
        {count}<span style={{ fontSize: 12, opacity: 0.7 }}> / {total}</span>
      </div>
    </div>
  );
};

/* Render simple markdown - headings, lists, paragraphs, bold, inline code */
const renderMarkdown = (text) => {
  if (!text) return null;
  const lines = text.split("\n");
  const out = [];
  let listBuf = [];
  let i = 0;

  const flushList = () => {
    if (listBuf.length) {
      out.push(<ul key={`ul-${out.length}`}>{listBuf.map((l, idx) => <li key={idx} dangerouslySetInnerHTML={{ __html: inlineMd(l) }} />)}</ul>);
      listBuf = [];
    }
  };
  const inlineMd = (s) => s
    .replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
    .replace(/\*([^*]+)\*/g, "<em>$1</em>");

  for (const line of lines) {
    if (/^### /.test(line)) { flushList(); out.push(<h3 key={i++}>{line.replace(/^### /, "")}</h3>); }
    else if (/^## /.test(line)) { flushList(); out.push(<h2 key={i++}>{line.replace(/^## /, "")}</h2>); }
    else if (/^- /.test(line)) { listBuf.push(line.replace(/^- /, "")); }
    else if (line.trim() === "") { flushList(); }
    else { flushList(); out.push(<p key={i++} dangerouslySetInnerHTML={{ __html: inlineMd(line) }} />); }
  }
  flushList();
  return <div className="md">{out}</div>;
};

window.Icon = Icon;
window.Pill = Pill;
window.Badge = Badge;
window.MetricCard = MetricCard;
window.SecHeading = SecHeading;
window.Collapsible = Collapsible;
window.ValidationTile = ValidationTile;
window.renderMarkdown = renderMarkdown;
