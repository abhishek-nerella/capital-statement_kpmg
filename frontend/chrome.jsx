/* Chrome: top header, sidebar, tab bar */

const KpmgLogo = ({ size = 22 }) => (
  <img src="assets/kpmg-logo.png" alt="KPMG"
    style={{
      height: size,
      width: "auto",
      display: "block",
      filter: "brightness(0) invert(1)",
    }} />
);

const TopHeader = ({ runId, asOf, darkNavyHero, onOpenAudit, onLogout }) => (
  <div>
    <div className="top-header" style={darkNavyHero ? { background: "var(--dark-navy)" } : {}}>
      <KpmgLogo size={20} />
      <div style={{ width: 1, height: 24, background: "rgba(255,255,255,0.2)" }}></div>
      <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.15 }}>
        <div style={{ fontSize: 13, fontWeight: 700, letterSpacing: "-0.005em" }}>
          Capital Analysis Statement Generator
        </div>
        <div style={{ fontSize: 10.5, opacity: 0.8, letterSpacing: "0.04em", textTransform: "uppercase" }}>
          Funds Reporting · Alternatives
        </div>
      </div>
      <div style={{ flex: 1 }}></div>
      {runId && (
        <div style={{ display: "flex", alignItems: "center", gap: 16, fontSize: 11.5 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, opacity: 0.85 }}>
            <Icon name="shield" size={13} />
            <span>Run <span style={{ fontFamily: "var(--font-mono)" }}>{runId}</span></span>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6, opacity: 0.85 }}>
            <Icon name="calendar" size={13} />
            <span>As of {asOf}</span>
          </div>
        </div>
      )}
      <button className="icon-btn dark" title="Audit log" onClick={onOpenAudit}><Icon name="audit" size={14} /></button>
      <button className="icon-btn dark" title="Settings"><Icon name="settings" size={14} /></button>
      <div style={{
        background: "var(--pacific)", color: "var(--dark-navy)",
        padding: "5px 10px", fontWeight: 800, fontSize: 10.5, letterSpacing: "0.08em",
      }}>INTERNAL · CONFIDENTIAL</div>
      {onLogout && (
        <button
          onClick={onLogout}
          title="Sign out"
          style={{
            display: "flex", alignItems: "center", gap: 6,
            padding: "5px 12px",
            background: "rgba(255,255,255,0.10)",
            border: "1px solid rgba(255,255,255,0.20)",
            color: "#fff",
            fontSize: 11.5,
            fontWeight: 600,
            cursor: "pointer",
            letterSpacing: "0.03em",
            marginLeft: 4,
          }}
          onMouseEnter={e => e.currentTarget.style.background = "rgba(255,255,255,0.20)"}
          onMouseLeave={e => e.currentTarget.style.background = "rgba(255,255,255,0.10)"}
        >
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none"
               stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
            <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/>
            <polyline points="16 17 21 12 16 7"/>
            <line x1="21" y1="12" x2="9" y2="12"/>
          </svg>
          Sign out
        </button>
      )}
    </div>
    <div className="accent-strip"></div>
  </div>
);

const Sidebar = ({
  fileState, fileName, onUpload, onFileSelected, onClearFile,
  scope, setScope, allInvestors, selectedInvestors, setSelectedInvestors,
  onOpenChat, activeTab,
}) => {
  const isPE = activeTab === "pe";
  const fileLoaded = fileState === "loaded";
  const fileLoading = fileState === "loading";
  const fileError = fileState === "error";

  const fileInputRef = React.useRef(null);
  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    onFileSelected && onFileSelected(file);
    e.target.value = ""; // allow re-selecting same file
  };

  return (
    <aside className="sidebar">
      {onFileSelected && (
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx"
          style={{ display: "none" }}
          onChange={handleFileChange}
        />
      )}

      {/* Brand strip (Apex sidebar nav header) */}
      <div style={{
        display: "flex", alignItems: "center", justifyContent: "space-between",
        padding: "14px 18px", borderBottom: "1px solid rgba(255,255,255,0.06)",
      }}>
        <KpmgLogo size={18} />
        <div style={{ transform: "rotate(180deg)", display: "flex" }}>
          <Icon name="chevron-right" size={13} color="rgba(255,255,255,0.50)" />
        </div>
      </div>

      {/* File upload */}
      <div className="side-section">
        <div className="side-section-head">
          <div className="side-tile"><Icon name="upload" size={13} /></div>
          <h4>File Upload</h4>
        </div>
        {!fileLoaded ? (
          <label
            className={`upload-zone ${fileLoading ? "active" : ""}`}
            style={{ display: "block" }}
            onClick={(e) => {
              e.preventDefault();
              if (onFileSelected && fileInputRef.current) {
                fileInputRef.current.click();
              } else if (onUpload) {
                onUpload();
              }
            }}>
            {fileLoading ? (
              <>
                <Icon name="refresh" size={22} color="var(--pacific)" />
                <div style={{ marginTop: 8, fontWeight: 700, color: "var(--white)" }}>Parsing…</div>
                <div style={{ marginTop: 4 }}>Reading {isPE ? "xlsx" : "PCAP"} columns</div>
              </>
            ) : fileError ? (
              <>
                <Icon name="warn" size={22} color="var(--danger)" />
                <div style={{ marginTop: 8, fontWeight: 700, color: "var(--white)" }}>Upload failed</div>
                <div style={{ marginTop: 4, color: "rgba(255,255,255,0.70)" }}>Click to retry</div>
              </>
            ) : (
              <>
                <Icon name="upload" size={22} color="var(--pacific)" />
                <div style={{ marginTop: 8, fontWeight: 700, color: "var(--white)" }}>Drop {isPE ? ".xlsx" : "PCAP"} file</div>
                <div style={{ marginTop: 4 }}>or <span style={{ color: "var(--pacific)", fontWeight: 600 }}>browse</span></div>
                <div style={{ marginTop: 6, fontSize: 10, color: "rgba(255,255,255,0.50)" }}>{isPE ? "Investor data sheet" : "PCAP + optional CF Ledger"}</div>
              </>
            )}
          </label>
        ) : (
          <div style={{ background: "rgba(255,255,255,0.06)", border: "1px solid rgba(172,234,255,0.18)", padding: 10 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <Icon name="file" size={14} color="var(--pacific)" />
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 11.5, fontWeight: 600, color: "var(--white)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {fileName}
                </div>
                <div style={{ fontSize: 10, color: "rgba(255,255,255,0.70)" }}>{allInvestors.length} investors · 1.2 MB</div>
              </div>
              <button className="icon-btn dark" onClick={onClearFile} title="Remove file"><Icon name="x" size={12} /></button>
            </div>
          </div>
        )}
      </div>

      {/* Investor selection — only when loaded */}
      {fileLoaded && (
        <div className="side-section">
          <div className="side-section-head">
            <div className="side-tile"><Icon name="users" size={13} /></div>
            <h4>Investor Selection</h4>
          </div>
          <div
            className={`side-radio ${scope === "all" ? "active" : ""}`}
            onClick={() => setScope("all")}>
            <div className="dot"></div>
            <span>All investors ({allInvestors.length})</span>
          </div>
          <div
            className={`side-radio ${scope === "selected" ? "active" : ""}`}
            onClick={() => setScope("selected")}>
            <div className="dot"></div>
            <span>Selected ({selectedInvestors.length})</span>
          </div>

          {scope === "selected" && (
            <div style={{ marginTop: 10 }}>
              <div style={{ position: "relative" }}>
                <input
                  type="text"
                  placeholder="Search investors…"
                  style={{
                    width: "100%", background: "rgba(255,255,255,0.06)",
                    border: "1px solid rgba(172,234,255,0.18)",
                    color: "var(--white)", fontSize: 11.5,
                    padding: "6px 8px 6px 28px"
                  }}
                />
                <span style={{ position: "absolute", left: 8, top: "50%", transform: "translateY(-50%)", color: "rgba(255,255,255,0.50)" }}>
                  <Icon name="search" size={12} />
                </span>
              </div>
              <div style={{
                marginTop: 8,
                maxHeight: 200, overflowY: "auto",
                background: "rgba(0,0,0,0.20)",
              }}>
                {allInvestors.slice(0, 12).map((inv) => {
                  const sel = selectedInvestors.includes(inv);
                  return (
                    <div
                      key={inv}
                      onClick={() => {
                        setSelectedInvestors(
                          sel ? selectedInvestors.filter(i => i !== inv) : [...selectedInvestors, inv]
                        );
                      }}
                      style={{
                        padding: "5px 10px", fontSize: 11.5, cursor: "pointer",
                        background: sel ? "rgba(0,184,245,0.15)" : "transparent",
                        color: "var(--white)",
                        borderLeft: sel ? "2px solid var(--pacific)" : "2px solid transparent",
                        display: "flex", alignItems: "center", gap: 6,
                      }}>
                      <Icon name={sel ? "check" : "plus"} size={11} color={sel ? "var(--pacific)" : "rgba(255,255,255,0.50)"} />
                      {inv}
                    </div>
                  );
                })}
                {allInvestors.length > 12 && (
                  <div style={{ padding: "6px 10px", fontSize: 10.5, color: "rgba(255,255,255,0.60)", textAlign: "center" }}>
                    +{allInvestors.length - 12} more
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      )}

      {/* Chat */}
      <div className="side-section">
        <div className="side-section-head">
          <div className="side-tile"><Icon name="msg" size={13} /></div>
          <h4>{isPE ? "PE Chat" : "HF Chat"}</h4>
        </div>
        <button
          onClick={onOpenChat}
          style={{
            width: "100%", display: "flex", alignItems: "center", gap: 8,
            padding: "10px 14px", fontSize: 12.5, fontWeight: 600,
            background: "var(--cobalt)", color: "var(--white)",
            border: "none", borderRadius: 8, cursor: "pointer",
            boxShadow: "0px 3px 6px rgba(0,0,0,0.10)",
            transition: "background 120ms",
          }}
          onMouseEnter={e => e.currentTarget.style.background = "var(--kpmg-blue)"}
          onMouseLeave={e => e.currentTarget.style.background = "var(--cobalt)"}
        >
          <Icon name="msg" size={14} />
          <span>Open {isPE ? "PE" : "HF"} Chat</span>
          <div style={{ flex: 1 }}></div>
          <Icon name="chevron-right" size={12} />
        </button>
      </div>

      <div style={{ flex: 1 }}></div>

      {/* Brand block at bottom */}
      <div className="side-section" style={{ borderBottom: "none", paddingTop: 12, paddingBottom: 16 }}>
        <KpmgLogo size={22} />
        <div style={{ marginTop: 10, fontSize: 10.5, color: "rgba(255,255,255,0.70)", letterSpacing: "0.08em", textTransform: "uppercase" }}>
          Capital Statements
        </div>
      </div>
    </aside>
  );
};

const TabBar = ({ activeTab, setTab, peCount, hfCount, runStats }) => (
  <div className="tabbar">
    <div className={`tab ${activeTab === "pe" ? "active" : ""}`} onClick={() => setTab("pe")}>
      <Icon name="users" size={14} />
      <span>Private Equity</span>
      {peCount > 0 && <span className="counter">{peCount}</span>}
    </div>
    <div className={`tab ${activeTab === "hf" ? "active" : ""}`} onClick={() => setTab("hf")}>
      <Icon name="bolt" size={14} />
      <span>Hedge Fund</span>
      {hfCount > 0 && <span className="counter">{hfCount}</span>}
    </div>
    <div style={{ flex: 1 }}></div>
    {runStats && (
      <div style={{ display: "flex", alignItems: "center", gap: 14, fontSize: 11.5, color: "var(--ink-500)", padding: "0 8px", whiteSpace: "nowrap" }}>
        <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <Icon name="dot" size={10} color="var(--teal)" />
          Live
        </span>
        {runStats.lastRun && (
          <span>Last run: <span style={{ color: "var(--ink-900)", fontFamily: "var(--font-mono)" }}>{runStats.lastRun}</span></span>
        )}
      </div>
    )}
  </div>
);

window.TopHeader = TopHeader;
window.Sidebar = Sidebar;
window.TabBar = TabBar;
window.KpmgLogo = KpmgLogo;
