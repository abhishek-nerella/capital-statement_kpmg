/* Chat drawer + Audit log drawer */

const ChatDrawer = ({ open, onClose, activeTab, sessionToken, apiBase }) => {
  const isPE = activeTab === "pe";
  const [messages, setMessages] = React.useState([
    {
      role: "ai",
      content: isPE
        ? "I'm your senior PE advisor. I've analyzed the uploaded portfolio data — feel free to ask about any specific investor, fund performance metrics, or fee structures."
        : "I'm your senior HF analyst. I've analyzed the fund's PCAP and waterfall data — feel free to ask about NAV attribution, IRR benchmarks, or redemption risks.",
    },
  ]);
  const [input, setInput] = React.useState("");
  const [thinking, setThinking] = React.useState(false);
  const scrollRef = React.useRef(null);

  React.useEffect(() => {
    if (scrollRef.current) scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
  }, [messages, thinking]);

  const send = async () => {
    if (!input.trim()) return;
    const userMsg = { role: "user", content: input.trim() };
    const updatedMessages = [...messages, userMsg];
    setMessages(updatedMessages);
    setInput("");
    setThinking(true);

    // ── Real API path ──────────────────────────────────────────────────────
    if (sessionToken && apiBase != null) {
      try {
        const endpoint = isPE ? `${apiBase}/api/pe/chat` : `${apiBase}/api/hf/chat`;
        const tokenKey = isPE ? "session_token" : "pcap_token";

        const res = await fetch(endpoint, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            [tokenKey]: sessionToken,
            // Map internal "ai" role to "assistant" for Gemini API
            messages: updatedMessages.map(m => ({
              role:    m.role === "ai" ? "assistant" : "user",
              content: m.content,
            })),
          }),
        });
        const data = await res.json();
        if (data.ok) {
          setMessages(m => [...m, { role: "ai", content: data.reply }]);
          setThinking(false);
          return;
        }
      } catch (err) {
        console.error("Chat API error:", err);
        // Fall through to mock response
      }
    }

    const errMsg = (sessionToken && apiBase != null)
      ? "Backend connection error — check that the server is running and try again."
      : "Not connected to backend. Upload a data file to enable live analysis.";
    setMessages(m => [...m, { role: "ai", content: `⚠️ ${errMsg}` }]);
    setThinking(false);
  };

  if (!open) return null;

  return (
    <>
      <div className="drawer-mask" onClick={onClose}></div>
      <div className="drawer">
        <div className="drawer-head">
          <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.2 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <Icon name="spark" size={14} />
              <span style={{ fontWeight: 700, fontSize: 13 }}>{isPE ? "PE Chat" : "HF Chat"}</span>
              <span style={{
                background: "var(--pacific)", color: "var(--dark-navy)",
                fontSize: 9.5, padding: "2px 6px", fontWeight: 800, letterSpacing: "0.08em"
              }}>GEMINI 2.5 PRO</span>
              {sessionToken && (
                <span style={{ fontSize: 9, color: "var(--teal)", fontWeight: 800, letterSpacing: "0.08em" }}>● LIVE</span>
              )}
            </div>
            <div style={{ fontSize: 10.5, opacity: 0.75, marginTop: 2 }}>
              Senior {isPE ? "PE advisor" : "HF analyst"}
            </div>
          </div>
          <div style={{ flex: 1 }}></div>
          <button className="icon-btn dark" title="New chat" onClick={() => setMessages(msgs => msgs.slice(0, 1))}><Icon name="plus" size={12} /></button>
          <button className="icon-btn dark" onClick={onClose} title="Close"><Icon name="x" size={12} /></button>
        </div>

        {/* Context strip */}
        <div style={{
          padding: "10px 14px",
          background: "var(--surface-2)",
          borderBottom: "1px solid var(--border)",
          display: "flex", alignItems: "center", gap: 8,
          fontSize: 11, color: "var(--ink-700)",
        }}>
          <Icon name="info" size={11} color="var(--cobalt)" />
          <span style={{ fontWeight: 600 }}>Context:</span>
          <span>
            {sessionToken
              ? (isPE ? "PE portfolio · connected" : "HF fund · connected")
              : "No data loaded — upload a file to begin"}
          </span>
          {!sessionToken && (
            <span style={{ marginLeft: "auto", fontSize: 10, color: "var(--purple)", fontStyle: "italic", fontWeight: 700 }}>
              not connected
            </span>
          )}
        </div>

        {/* Messages */}
        <div ref={scrollRef} style={{ flex: 1, overflow: "auto", padding: 16, background: "var(--white)" }}>
          {messages.map((m, i) => (
            <div key={i} className={`bubble ${m.role}`}>
              <div className="b-role">
                {m.role === "user" ? (
                  <><Icon name="user" size={10} /> You</>
                ) : (
                  <><Icon name="spark" size={10} color="var(--pacific)" /> {isPE ? "Gemini · PE Advisor" : "Gemini · HF Analyst"}</>
                )}
              </div>
              <div style={{ color: "var(--ink-900)" }}>
                {m.role === "ai" ? renderMarkdown(m.content) : m.content}
              </div>
            </div>
          ))}
          {thinking && (
            <div className="bubble ai">
              <div className="b-role"><Icon name="spark" size={10} color="var(--pacific)" /> {isPE ? "Gemini · PE Advisor" : "Gemini · HF Analyst"}</div>
              <div style={{ display: "flex", alignItems: "center", gap: 6, color: "var(--ink-500)", fontSize: 12 }}>
                <span className="pulse">●</span>
                <span className="pulse" style={{ animationDelay: "0.2s" }}>●</span>
                <span className="pulse" style={{ animationDelay: "0.4s" }}>●</span>
                <span style={{ marginLeft: 8 }}>Thinking…</span>
              </div>
            </div>
          )}
        </div>

        {/* Suggestion chips */}
        <div style={{ padding: "10px 14px", borderTop: "1px solid var(--border)", display: "flex", flexWrap: "wrap", gap: 6 }}>
          {(isPE ? [
            "Top 5 LPs by TVPI",
            "Pacing for Q3 capital call",
            "Why is Investor J failing?",
            "Exit scenarios for top 4 positions",
          ] : [
            "Stress at -30% NAV",
            "GP carry by LP",
            "Unit price decomp",
            "HWM crossovers next quarter",
          ]).map(s => (
            <button key={s} onClick={() => setInput(s)} style={{
              fontSize: 11, padding: "4px 10px", border: "1px solid var(--border)",
              background: "var(--white)", cursor: "pointer", color: "var(--ink-700)",
              fontFamily: "inherit"
            }}>{s}</button>
          ))}
        </div>

        {/* Input */}
        <div style={{ padding: 14, borderTop: "1px solid var(--border)" }}>
          <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
            <textarea
              value={input}
              onChange={e => setInput(e.target.value)}
              onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } }}
              placeholder={`Ask anything about the ${isPE ? "portfolio" : "fund"}…`}
              style={{ flex: 1, resize: "none", height: 60, fontSize: 13 }}
            />
            <button className="btn btn-primary" onClick={send} disabled={!input.trim() || thinking}>
              <Icon name="send" size={13} />
              Send
            </button>
          </div>
          <div style={{ marginTop: 6, fontSize: 10.5, color: "var(--ink-500)", display: "flex", justifyContent: "space-between" }}>
            <span>↵ to send · ⇧↵ for newline</span>
            <span>{messages.length} messages · ~3.2K tokens</span>
          </div>
        </div>
      </div>
    </>
  );
};

const AuditDrawer = ({ open, onClose, apiBase }) => {
  const [view, setView] = React.useState("jsonl"); // "jsonl" | "plain"
  const [entries, setEntries] = React.useState([]);
  const [plainLines, setPlainLines] = React.useState([]);

  React.useEffect(() => {
    if (!open) return;
    if (apiBase == null) { setEntries([]); setPlainLines([]); return; }
    fetch(`${apiBase}/api/audit/entries`)
      .then(r => r.json())
      .then(d => setEntries(Array.isArray(d) ? d : (d.entries || [])))
      .catch(() => setEntries([]));
    fetch(`${apiBase}/api/audit/plain`)
      .then(r => r.json())
      .then(d => setPlainLines(Array.isArray(d) ? d : (d.lines || [])))
      .catch(() => setPlainLines([]));
  }, [open, apiBase]);

  const handleDownload = async () => {
    if (apiBase == null) return;
    const isPlain = view === "plain";
    try {
      const res = await fetch(`${apiBase}/api/audit/${isPlain ? "download-plain" : "download"}`);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = isPlain ? "audit_log.txt" : "audit_log.jsonl";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Audit download failed:", err);
    }
  };

  if (!open) return null;
  const count = view === "plain" ? plainLines.length : entries.length;
  return (
    <>
      <div className="drawer-mask" onClick={onClose}></div>
      <div className="drawer" style={{ width: 560 }}>
        <div className="drawer-head">
          <Icon name="audit" size={14} />
          <span style={{ fontWeight: 700, fontSize: 13 }}>Audit Trail</span>
          <div style={{ display: "flex", gap: 4, marginLeft: 4 }}>
            <button
              onClick={() => setView("jsonl")}
              style={{
                background: view === "jsonl" ? "var(--pacific)" : "rgba(255,255,255,0.10)",
                color: view === "jsonl" ? "var(--dark-navy)" : "#fff",
                border: "none", cursor: "pointer",
                fontSize: 9.5, padding: "2px 8px", fontWeight: 800, letterSpacing: "0.08em",
                borderRadius: 4,
              }}
            >JSONL</button>
            <button
              onClick={() => setView("plain")}
              style={{
                background: view === "plain" ? "var(--pacific)" : "rgba(255,255,255,0.10)",
                color: view === "plain" ? "var(--dark-navy)" : "#fff",
                border: "none", cursor: "pointer",
                fontSize: 9.5, padding: "2px 8px", fontWeight: 800, letterSpacing: "0.08em",
                borderRadius: 4,
              }}
            >PLAIN ENGLISH</button>
          </div>
          <div style={{ flex: 1 }}></div>
          <button
            className="icon-btn dark"
            title={view === "plain" ? "Download audit_log.txt" : "Download audit_log.jsonl"}
            onClick={handleDownload}
            disabled={apiBase == null}
          ><Icon name="download" size={12} /></button>
          <button className="icon-btn dark" onClick={onClose}><Icon name="x" size={12} /></button>
        </div>
        <div style={{ padding: "10px 14px", background: "var(--surface-2)", borderBottom: "1px solid var(--border)", fontSize: 11, color: "var(--ink-700)" }}>
          <span style={{ fontFamily: "var(--font-mono)" }}>{view === "plain" ? "audit_log.txt" : "audit_log.jsonl"}</span> · <span style={{ fontFamily: "var(--font-mono)", color: "var(--ink-900)", fontWeight: 700 }}>{count} event{count !== 1 ? "s" : ""}</span>
          {apiBase != null && (
            <span style={{ marginLeft: 12, color: "var(--teal)", fontWeight: 700 }}>● live file on server</span>
          )}
        </div>
        <div style={{ flex: 1, overflow: "auto", padding: 12, background: "#0c1320", fontFamily: "var(--font-mono)", fontSize: 11, color: "#d8e2f1", lineHeight: 1.6 }}>
          {view === "plain" ? (
            plainLines.length > 0
              ? plainLines.map((line, i) => <PlainAuditLine key={i} line={line} />)
              : (
                <div style={{ padding: "30px 16px", textAlign: "center", color: "#4a5a7a", fontSize: 12 }}>
                  {apiBase != null ? "No audit events yet — generate statements to populate the log." : "Connect to backend to view audit events."}
                </div>
              )
          ) : (
            entries.length > 0
              ? entries.map((e, i) => <AuditLine key={i} entry={e} />)
              : (
                <div style={{ padding: "30px 16px", textAlign: "center", color: "#4a5a7a", fontSize: 12 }}>
                  {apiBase != null ? "No audit events yet — generate statements to populate the log." : "Connect to backend to view audit events."}
                </div>
              )
          )}
        </div>
      </div>
    </>
  );
};

const PlainAuditLine = ({ line }) => {
  const isFail = /FAILED/.test(line);
  const isStart = /Run .* started/.test(line);
  const isComplete = /Run .* completed/.test(line);
  const color = isFail ? "#ff9ad1" : (isStart || isComplete) ? "#7ad6ff" : "#d8e2f1";
  return (
    <div style={{ marginBottom: 6, paddingBottom: 6, borderBottom: "1px solid #1a2540", color }}>
      {line}
    </div>
  );
};

const AuditLine = ({ entry }) => {
  const colorMap = {
    wrangler_change:     "#7fd1d0",
    validation_pass:     "#7adbab",
    validation_revalued: "#7ad6ff",
    validation_fail:     "#ff9ad1",
    document_failed:     "#ff9ad1",
    document_generated:  "#cbb6ff",
    gemini_insight:      "#cbb6ff",
  };
  const c = colorMap[entry.type] || "#d8e2f1";
  return (
    <div style={{ marginBottom: 8, paddingBottom: 8, borderBottom: "1px solid #1a2540" }}>
      <span style={{ color: "#7f8da8" }}>"{entry.ts}"</span>{" "}
      <span style={{ color: c, fontWeight: 700 }}>{entry.type}</span>{" "}
      <span style={{ color: "#fff" }}>{entry.investor}</span>
      <div style={{ paddingLeft: 14, color: "#a9b6cc", marginTop: 2 }}>↳ {entry.detail}</div>
    </div>
  );
};

window.ChatDrawer = ChatDrawer;
window.AuditDrawer = AuditDrawer;
