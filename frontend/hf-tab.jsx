/* Hedge Fund tab */

const HFTab = ({
  fileState, selectedInvestors, scope,
  generationState, results, onTriggerGenerate,
  insightState, insightMd, insightType, setInsightType,
  insightScope, setInsightScope, insightInvestor, setInsightInvestor,
  onTriggerInsight,
  hasLedger, onUploadLedger, onLedgerSelected,
  darkNavyHero,
  downloadUrls,
  runId,
  pcapPreviewData,
  fundStats,
  generationError,
  insightError,
  pcapFileName,
  ledgerFileName,
}) => {
  if (fileState === "empty") return <HFEmpty />;
  if (fileState === "loading") return <HFLoading />;
  if (fileState === "error")  return <HFErrorState />;

  const displayInvestors = pcapPreviewData || [];
  const fundName = (displayInvestors.length > 0 &&
    (displayInvestors[0].FUND_NAME || displayInvestors[0].PARTNERSHIP_NAME)) || "Hedge Fund";

  const totalNav = fundStats ? (fundStats.total_nav_cq || 0)
    : displayInvestors.reduce((s, i) => s + (Number(i.END_CAP_CQ) || 0), 0);
  const avgIrr = fundStats ? (fundStats.avg_net_irr || 0)
    : displayInvestors.length > 0
      ? displayInvestors.reduce((s, i) => s + (Number(i.NET_IRR) || 0), 0) / displayInvestors.length
      : 0;
  const avgTvpi = fundStats ? (fundStats.avg_tvpi || 0)
    : displayInvestors.length > 0
      ? displayInvestors.reduce((s, i) => s + (Number(i.TVPI) || 0), 0) / displayInvestors.length
      : 0;
  const totalCarry = displayInvestors.reduce((s, i) => {
    const excess = Number(i.EXCESS_HURDLE) || 0;
    const feeRate = Number(i.INC_FEE_RATE) || 0.20;
    return s + excess * feeRate;
  }, 0);

  const targetInvestors = scope === "all"
    ? displayInvestors
    : displayInvestors.filter(i => selectedInvestors.includes(i.INVESTOR_NAME));

  return (
    <div style={{ padding: "20px 24px 60px" }}>
      {/* PCAP + Ledger upload status */}
      <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
        <UploadStatusCard
          icon="file"
          label="PCAP Excel"
          status="loaded"
          fileName={pcapFileName || "pcap.xlsx"}
          meta={`${displayInvestors.length} investors · live data`} />
        <UploadStatusCard
          icon="audit"
          label="CF Ledger (optional)"
          status={hasLedger ? "loaded" : "empty"}
          fileName={hasLedger ? (ledgerFileName || "cf_ledger.xlsx") : null}
          meta={hasLedger ? "CF transactions loaded" : "Drag .xlsx or click to add"}
          onUpload={onUploadLedger}
          onFileSelected={onLedgerSelected} />
      </div>

      {/* Hero metric strip */}
      <div className={darkNavyHero ? "hero-darknav" : ""}
        style={{
          padding: darkNavyHero ? "20px 24px" : "0",
          marginBottom: 22,
          marginLeft: darkNavyHero ? -24 : 0,
          marginRight: darkNavyHero ? -24 : 0,
        }}>
        {darkNavyHero && (
          <div style={{ display: "flex", alignItems: "baseline", gap: 14, marginBottom: 14 }}>
            <h1 style={{ fontSize: 22, margin: 0, fontWeight: 800, letterSpacing: "-0.02em", color: "var(--white)" }}>
              {fundName}
            </h1>
            <span style={{ fontSize: 12, color: "var(--light-blue)" }}>Hedge Fund · PCAP · Live Data</span>
          </div>
        )}
        {!darkNavyHero && (
          <div style={{ display: "flex", alignItems: "baseline", gap: 14, marginBottom: 14, flexWrap: "wrap" }}>
            <h1 style={{ fontSize: 20, margin: 0, fontWeight: 800, letterSpacing: "-0.02em", color: "var(--kpmg-blue)" }}>
              {fundName}
            </h1>
            <span style={{ fontSize: 11.5, color: "var(--ink-500)", letterSpacing: "0.04em", textTransform: "uppercase" }}>
              Hedge Fund · PCAP Live
            </span>
          </div>
        )}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, minmax(0, 1fr))", gap: 10 }}>
          <MetricCard label="Limited Partners"
            value={<span className="num">{displayInvestors.length}</span>}
            sub={<span>{displayInvestors.length} investor{displayInvestors.length !== 1 ? "s" : ""}</span>}
            accent="cobalt" />
          <MetricCard label="Total NAV (CQ)"
            value={<span className="num">{window.fmtMoney(totalNav)}</span>}
            sub={<span>End capital quarter</span>}
            accent="pacific" />
          <MetricCard label="Avg Net IRR · TVPI"
            value={<span className="num">{avgIrr.toFixed(2)}% · {avgTvpi.toFixed(2)}x</span>}
            sub={<span>Hurdle 4.0%</span>}
            accent="teal" />
          <MetricCard label="GP Carry Accrued"
            value={<span className="num">{window.fmtMoney(totalCarry)}</span>}
            sub={<span>{displayInvestors.length} LPs in fund</span>}
            accent="pacific" />
        </div>
      </div>

      {/* Selected pills */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 18 }}>
        <span style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", color: "var(--ink-500)" }}>
          Selected ({targetInvestors.length})
        </span>
        {targetInvestors.slice(0, 5).map(inv => (
          <Pill key={inv.INVESTOR_NAME} color="blue">{inv.INVESTOR_NAME}</Pill>
        ))}
        {targetInvestors.length > 5 && <Pill>+{targetInvestors.length - 5} more</Pill>}
      </div>

      {/* PCAP Preview */}
      <SecHeading right={
        displayInvestors.length > 0 && (
          <span style={{ fontSize: 10, color: "var(--teal)", fontWeight: 700, letterSpacing: "0.08em" }}>
            ● LIVE DATA
          </span>
        )
      }>Preview PCAP Data</SecHeading>
      <div className="card card-pacific" style={{ marginBottom: 22 }}>
        <HFPcapGrid investors={displayInvestors} />
      </div>

      {/* Generate */}
      <SecHeading>Generate Documents</SecHeading>
      {generationError && (
        <div style={{ marginBottom: 12, padding: "10px 14px", background: "#fdf2f4", border: "1px solid var(--purple)", display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: "var(--purple)" }}>
          <Icon name="warn" size={13} color="var(--purple)" />
          <span style={{ flex: 1 }}>{generationError}</span>
        </div>
      )}
      <div className="card" style={{ padding: 16, marginBottom: 22 }}>
        <HFGeneratePanel
          targetInvestors={targetInvestors}
          generationState={generationState}
          results={results}
          onTriggerGenerate={onTriggerGenerate}
          runId={runId}
        />
      </div>

      {/* Downloads + companion */}
      {generationState === "done" && (
        <>
          <SecHeading>Download Generated Documents</SecHeading>
          <div className="card" style={{ padding: 16, marginBottom: 12 }}>
            <HFDownloads count={results.filter(r => r.ok).length} downloadUrls={downloadUrls} />
          </div>
          <CompanionExcelCard />
        </>
      )}

      {/* Insights */}
      <SecHeading right={
        <span style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 10.5, color: "var(--ink-500)", fontWeight: 600 }}>
          <Icon name="spark" size={11} color="var(--cobalt)" />
          Gemini 2.5 Pro · 65K tokens
        </span>
      }>AI HF Insights</SecHeading>

      <div className="card" style={{ padding: 16 }}>
        {insightError && (
          <div style={{ marginBottom: 12, padding: "10px 14px", background: "#fdf2f4", border: "1px solid var(--purple)", display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: "var(--purple)" }}>
            <Icon name="warn" size={13} color="var(--purple)" />
            <span style={{ flex: 1 }}>{insightError}</span>
          </div>
        )}
        <HFInsights
          insightState={insightState} insightMd={insightMd}
          insightType={insightType} setInsightType={setInsightType}
          insightScope={insightScope} setInsightScope={setInsightScope}
          insightInvestor={insightInvestor} setInsightInvestor={setInsightInvestor}
          onTriggerInsight={onTriggerInsight}
          allInvestors={displayInvestors.map(i => i.INVESTOR_NAME)}
        />
      </div>
    </div>
  );
};

const UploadStatusCard = ({ icon, label, status, fileName, meta, onUpload, onFileSelected }) => {
  const loaded = status === "loaded";
  const fileInputRef = React.useRef(null);

  const handleBtnClick = () => {
    if (onFileSelected && fileInputRef.current) {
      fileInputRef.current.click();
    } else if (onUpload) {
      onUpload();
    }
  };

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    onFileSelected && onFileSelected(file);
    e.target.value = "";
  };

  return (
    <div className="card" style={{ flex: 1, padding: "12px 14px", borderTop: `3px solid ${loaded ? "var(--teal)" : "var(--ink-200)"}` }}>
      {onFileSelected && (
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx"
          style={{ display: "none" }}
          onChange={handleFileChange}
        />
      )}
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <div style={{
          width: 36, height: 36, background: loaded ? "var(--teal)" : "var(--ink-100)",
          color: loaded ? "var(--white)" : "var(--ink-500)",
          display: "inline-flex", alignItems: "center", justifyContent: "center"
        }}>
          <Icon name={loaded ? "check" : icon} size={16} />
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: "var(--ink-500)" }}>{label}</span>
            {loaded ? (
              <span className="badge badge-pass" style={{ fontSize: 9 }}>LOADED</span>
            ) : (
              <span className="badge" style={{ fontSize: 9, background: "var(--ink-100)", color: "var(--ink-500)" }}>OPTIONAL</span>
            )}
          </div>
          <div style={{ fontSize: 13, fontWeight: 600, color: "var(--ink-900)", marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontFamily: loaded ? "var(--font-mono)" : "var(--font-sans)" }}>
            {fileName || "No file uploaded"}
          </div>
          <div style={{ fontSize: 11, color: "var(--ink-500)" }}>{meta}</div>
        </div>
        {!loaded && (onUpload || onFileSelected) && (
          <button className="btn btn-ghost btn-sm" onClick={handleBtnClick}>
            <Icon name="upload" size={11} /> Upload
          </button>
        )}
      </div>
    </div>
  );
};

const HFPcapGrid = ({ investors = [] }) => (
  <div style={{ maxHeight: 320, overflow: "auto" }}>
    <table className="dgrid">
      <thead>
        <tr>
          <th>Investor</th>
          <th>CCY</th>
          <th>Inception</th>
          <th className="right">Beg unit</th>
          <th className="right">End unit</th>
          <th className="right">Beg Cap CQ</th>
          <th className="right">End Cap CQ</th>
          <th className="right">Gross IRR</th>
          <th className="right">Net IRR</th>
          <th className="right">TVPI</th>
          <th className="right">LP Net Waterfall</th>
          <th>Lockup</th>
          <th>HWM</th>
        </tr>
      </thead>
      <tbody>
        {investors.map((inv, i) => {
          const name      = inv.INVESTOR_NAME  || "—";
          const ccy       = inv.REPT_CCY       || "USD";
          const inception = inv.INCEPTION_DATE || "—";
          const begPx     = Number(inv.BEG_PX)     || 0;
          const endPx     = Number(inv.END_PX)     || 0;
          const begCapCq  = Number(inv.BEG_CAP_CQ) || 0;
          const endCapCq  = Number(inv.END_CAP_CQ) || 0;
          const grossIrr  = Number(inv.GROSS_IRR)  || 0;
          const netIrr    = Number(inv.NET_IRR)    || 0;
          const tvpi      = Number(inv.TVPI)       || 0;
          const lpNetWf   = Number(inv.LP_NET_WF)  || 0;
          const lockupMo  = Number(inv.LOCKUP_MO)  || 0;
          const lockupExp = inv.LOCKUP_EXPIRED;
          const hwmActive = inv.HWM_ACTIVE;
          return (
            <tr key={name + i} className={i === 0 ? "selected" : ""}>
              <td style={{ fontWeight: 600 }}>{name}</td>
              <td>{ccy}</td>
              <td style={{ fontFamily: "var(--font-mono)", fontSize: 11, color: "var(--ink-500)" }}>
                {typeof inception === "string" ? inception.slice(0, 10) : "—"}
              </td>
              <td className="right">{begPx.toFixed(2)}</td>
              <td className="right">{endPx.toFixed(2)}</td>
              <td className="right">{window.fmtMoneyFull(begCapCq, ccy)}</td>
              <td className="right" style={{ fontWeight: 700 }}>{window.fmtMoneyFull(endCapCq, ccy)}</td>
              <td className="right">{grossIrr.toFixed(2)}%</td>
              <td className="right" style={{ color: netIrr > 12 ? "var(--teal)" : "var(--ink-900)" }}>{netIrr.toFixed(2)}%</td>
              <td className="right">{tvpi.toFixed(2)}x</td>
              <td className="right">{window.fmtMoneyFull(lpNetWf, ccy)}</td>
              <td>
                <span style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 11 }}>
                  <Icon name={lockupExp ? "check" : "calendar"} size={11}
                    color={lockupExp ? "var(--teal)" : "var(--ink-400)"} />
                  {lockupExp ? "Expired" : lockupMo ? `${lockupMo}mo` : "—"}
                </span>
              </td>
              <td>
                {hwmActive ? <Pill color="blue">HWM</Pill> : <span style={{ color: "var(--ink-400)" }}>—</span>}
              </td>
            </tr>
          );
        })}
      </tbody>
    </table>
  </div>
);

const HFGeneratePanel = ({ targetInvestors, generationState, results, onTriggerGenerate, runId }) => {
  if (generationState === "idle") {
    return (
      <div>
        <div style={{ display: "flex", alignItems: "flex-start", gap: 18, marginBottom: 8 }}>
          <div style={{ flex: 1 }}>
            <h3 style={{ margin: "0 0 6px", fontSize: 16, color: "var(--ink-900)" }}>
              Generate {targetInvestors.length} HF capital account statement{targetInvestors.length === 1 ? "" : "s"}
            </h3>
            <p style={{ margin: 0, fontSize: 12.5, color: "var(--ink-500)", maxWidth: 620 }}>
              Each LP yields a Word + PDF statement plus a companion <code style={{ background: "var(--ink-100)", padding: "1px 4px" }}>hf_pcap_model.xlsx</code> with live-formula PCAP, Waterfall, Cashflow IRR and Stress IRR sheets.
            </p>
            <div style={{ marginTop: 12, display: "flex", gap: 10, fontSize: 11, flexWrap: "wrap" }}>
              {["Period_Params", "PCAP", "Capital_Accounts", "CF_Aggregator", "Waterfall", "Cashflow_IRR", "Stress_IRR", "Dashboard"].map(s => (
                <span key={s} style={{ padding: "3px 8px", border: "1px solid var(--border)", color: "var(--ink-700)", background: "var(--white)" }}>{s}</span>
              ))}
            </div>
          </div>
          <button className="btn btn-primary btn-lg" onClick={onTriggerGenerate}>
            <Icon name="play" size={14} />
            Generate {targetInvestors.length} HF Statement{targetInvestors.length === 1 ? "" : "s"}
          </button>
        </div>
      </div>
    );
  }
  if (generationState === "running") {
    const progress = (results?.length || 0) / targetInvestors.length;
    return (
      <div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
          <Icon name="refresh" size={16} color="var(--cobalt)" />
          <span style={{ fontWeight: 700 }}>Building PCAP rolls and statements…</span>
          <span style={{ marginLeft: "auto", fontSize: 12, color: "var(--ink-500)", fontFamily: "var(--font-mono)" }}>
            {results?.length || 0} / {targetInvestors.length}
          </span>
        </div>
        <div className="progress-track"><div className="progress-fill" style={{ width: `${progress * 100}%` }}></div></div>
        <div style={{ marginTop: 14, maxHeight: 220, overflow: "auto" }}>
          {(results || []).slice().reverse().map((r, i) => (
            <div key={i} className="result-row">
              <Icon name="check" size={14} color="var(--teal)" stroke={2.5} />
              <div>
                <div className="r-name">{r.investor}</div>
                <div className="r-file">{r.file}</div>
              </div>
              <span style={{ fontSize: 11, color: "var(--ink-500)" }}>docx · pdf · xlsx</span>
              <div></div>
            </div>
          ))}
        </div>
      </div>
    );
  }
  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
        <div style={{ background: "var(--teal)", color: "var(--white)", padding: 10 }}>
          <Icon name="check" size={18} />
        </div>
        <div>
          <div style={{ fontSize: 14, fontWeight: 700 }}>
            Generated {results.length} HF statements + companion workbook
          </div>
          <div style={{ fontSize: 12, color: "var(--ink-500)" }}>
            Run · <span style={{ fontFamily: "var(--font-mono)" }}>{runId ? runId.slice(0, 8) : "—"}</span> · PCAP model built with live SUMIFS
          </div>
        </div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 6 }}>
          <button className="btn btn-ghost btn-sm"><Icon name="refresh" size={12} /> Re-run</button>
        </div>
      </div>
    </div>
  );
};

const HFDownloads = ({ count, downloadUrls }) => (
  <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10 }}>
    {[
      { i: "doc",   t: "Word ZIP",         s: `${count} HF .docx`,          cls: "btn-accent", urlKey: "word_zip" },
      { i: "file",  t: "PDF ZIP",          s: `${count} HF .pdf`,           cls: "btn-accent", urlKey: "pdf_zip" },
      { i: "table", t: "Summary Excel",    s: "Per-investor summary",       cls: "btn-ghost",  urlKey: "summary_excel" },
      { i: "bolt",  t: "PCAP Model .xlsx", s: "10 sheets · live formulas",  cls: "btn-dark",   urlKey: "companion_xlsx" },
    ].map(({ i, t, s, cls, urlKey }) => {
      const url = downloadUrls?.[urlKey];
      const btn = (
        <button
          className={`btn ${cls}`}
          style={{ width: "100%", padding: 14, flexDirection: "column", alignItems: "flex-start", gap: 6, whiteSpace: "normal", textAlign: "left", opacity: url ? 1 : 0.62 }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 6, width: "100%" }}>
            <Icon name={i} size={14} />
            <span style={{ fontSize: 13, fontWeight: 700 }}>{t}</span>
            <Icon name="download" size={11} style={{ marginLeft: "auto" }} />
          </div>
          <div style={{ fontSize: 11, fontWeight: 500, opacity: 0.85 }}>{s}</div>
        </button>
      );
      return url ? (
        <a key={t} href={url} download style={{ textDecoration: "none" }}>{btn}</a>
      ) : (
        <div key={t}>{btn}</div>
      );
    })}
  </div>
);

const CompanionExcelCard = () => (
  <div className="card" style={{ marginBottom: 22, borderTop: "3px solid var(--cobalt)" }}>
    <div style={{ padding: "12px 16px", display: "flex", alignItems: "center", gap: 10, borderBottom: "1px solid var(--border)" }}>
      <Icon name="table" size={14} color="var(--cobalt)" />
      <strong style={{ fontSize: 13, color: "var(--ink-900)" }}>hf_pcap_model.xlsx</strong>
      <span className="badge badge-reval" style={{ fontSize: 9 }}>10 SHEETS</span>
      <span style={{ marginLeft: "auto", fontSize: 11, color: "var(--ink-500)" }}>Live SUMIFS · stress models</span>
    </div>
    <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)" }}>
      {[
        { name: "Period_Params",          desc: "Quarter & rates config",    i: "calendar" },
        { name: "Dashboard",              desc: "Summary KPIs",              i: "chart" },
        { name: "Investor_Register",      desc: "LP master list",            i: "users" },
        { name: "PCAP",                   desc: "Full PCAP computation",     i: "table" },
        { name: "Capital_Accounts",       desc: "Per-investor roll",         i: "refresh" },
        { name: "CF_Ledger",              desc: "Raw transactions",          i: "audit" },
        { name: "CF_Aggregator",          desc: "Pivot · live SUMIFS",       i: "table" },
        { name: "Distribution_Waterfall", desc: "GP / LP split",             i: "bolt" },
        { name: "Cashflow_IRR",           desc: "IRR · XIRR formulas",       i: "chart" },
        { name: "Stress_IRR",             desc: "−10/−20/−30% NAV haircuts", i: "warn" },
      ].map((s, i) => (
        <div key={s.name} style={{
          padding: "10px 14px",
          borderRight: i % 5 !== 4 ? "1px solid var(--border)" : "none",
          borderBottom: i < 5 ? "1px solid var(--border)" : "none",
          display: "flex", flexDirection: "column", gap: 4
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <Icon name={s.i} size={11} color="var(--cobalt)" />
            <span style={{ fontFamily: "var(--font-mono)", fontSize: 10.5, fontWeight: 700, color: "var(--ink-900)" }}>{s.name}</span>
          </div>
          <div style={{ fontSize: 11, color: "var(--ink-500)" }}>{s.desc}</div>
        </div>
      ))}
    </div>
  </div>
);

const HFInsights = ({
  insightState, insightMd, insightType, setInsightType,
  insightScope, setInsightScope, insightInvestor, setInsightInvestor,
  onTriggerInsight, allInvestors,
}) => (
  <div>
    <div style={{ display: "grid", gridTemplateColumns: "1fr 1.4fr", gap: 16, marginBottom: 16 }}>
      <div>
        <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: "var(--ink-500)", marginBottom: 6 }}>Scope</div>
        <div style={{ display: "flex", border: "1px solid var(--border-strong)" }}>
          <button onClick={() => setInsightScope("portfolio")}
            style={{ flex: 1, padding: "10px", border: "none", background: insightScope === "portfolio" ? "var(--kpmg-blue)" : "var(--white)", color: insightScope === "portfolio" ? "var(--white)" : "var(--ink-700)", fontWeight: 600, fontSize: 12, cursor: "pointer" }}>
            <Icon name="users" size={12} /> &nbsp;Portfolio
          </button>
          <button onClick={() => setInsightScope("investor")}
            style={{ flex: 1, padding: "10px", border: "none", background: insightScope === "investor" ? "var(--kpmg-blue)" : "var(--white)", color: insightScope === "investor" ? "var(--white)" : "var(--ink-700)", fontWeight: 600, fontSize: 12, cursor: "pointer" }}>
            <Icon name="user" size={12} /> &nbsp;Individual
          </button>
        </div>
        {insightScope === "investor" && (
          <select value={insightInvestor} onChange={e => setInsightInvestor(e.target.value)} style={{ width: "100%", marginTop: 8 }}>
            {(allInvestors || []).map(name => (
              <option key={name}>{name}</option>
            ))}
          </select>
        )}
      </div>
      <div>
        <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: "var(--ink-500)", marginBottom: 6 }}>Insight Type</div>
        <div className="chip-select">
          {window.HF_INSIGHT_TYPES.map(t => (
            <button key={t.key} className={`chip ${insightType === t.key ? "active" : ""}`} onClick={() => setInsightType(t.key)}>
              {t.label}
            </button>
          ))}
        </div>
      </div>
    </div>

    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
      <button className="btn btn-dark btn-lg" onClick={onTriggerInsight} disabled={insightState === "running"}>
        <Icon name="spark" size={14} />
        {insightState === "running" ? "Analyzing…" : "Generate HF Insights"}
      </button>
      <span style={{ fontSize: 11, color: "var(--ink-500)" }}>
        Senior HF analyst persona · Citadel · Bridgewater · DE Shaw · Millennium
      </span>
    </div>

    {insightState === "running" && <InsightLoading />}
    {insightState === "done" && (
      <div style={{
        background: "linear-gradient(180deg, rgba(0,184,245,0.05), transparent 40%)",
        border: "1px solid var(--border)", borderTop: "3px solid var(--cobalt)",
        padding: "16px 20px",
      }}>
        {renderMarkdown(insightMd)}
        <div style={{ marginTop: 16, paddingTop: 12, borderTop: "1px dashed var(--border)", display: "flex", alignItems: "center", gap: 10, fontSize: 11, color: "var(--ink-500)" }}>
          <Icon name="info" size={11} />
          <span>Persona: Senior HF Analyst · Gemini 2.5 Pro · 65K tokens</span>
          <span style={{ marginLeft: "auto" }}>Logged to audit trail</span>
        </div>
      </div>
    )}
    {insightState === "idle" && (
      <div style={{ padding: "30px 16px", textAlign: "center", color: "var(--ink-500)", fontSize: 12.5, background: "var(--surface-2)", border: "1px dashed var(--border)" }}>
        <Icon name="spark" size={18} color="var(--ink-300)" />
        <div style={{ marginTop: 6 }}>Pick a scope + insight type, then generate.</div>
      </div>
    )}
  </div>
);

const HFEmpty = () => (
  <div style={{ padding: "60px 40px" }}>
    <div className="card card-pacific" style={{ padding: "48px 40px", textAlign: "center" }}>
      <div style={{ display: "inline-flex", padding: 18, border: "1.5px dashed var(--border-strong)", marginBottom: 18 }}>
        <Icon name="bolt" size={32} color="var(--kpmg-blue)" />
      </div>
      <h1 style={{ fontSize: 24, color: "var(--kpmg-blue)", margin: "0 0 8px", fontWeight: 800, letterSpacing: "-0.02em" }}>
        Upload PCAP Excel to begin
      </h1>
      <p style={{ color: "var(--ink-500)", fontSize: 13.5, maxWidth: 520, margin: "0 auto 18px" }}>
        Hedge Fund module reads a pre-calculated PCAP workbook (one row per LP, ~142 columns). Optionally include a CF Ledger of cashflows for a richer companion model.
      </p>
      <div style={{ display: "inline-flex", gap: 8 }}>
        <button className="btn btn-primary"><Icon name="upload" size={13} /> Upload PCAP</button>
        <button className="btn btn-ghost"><Icon name="upload" size={13} /> Add CF Ledger</button>
      </div>
      <div style={{ marginTop: 28, fontSize: 11, color: "var(--ink-500)" }}>
        <Icon name="info" size={11} /> &nbsp; <strong style={{ color: "var(--ink-700)" }}>Tip:</strong> Row 1 (merged group headers) is skipped automatically. Row 2 must contain canonical column names.
      </div>
    </div>
  </div>
);

const HFLoading = () => (
  <div style={{ padding: 30 }}>
    <div className="card" style={{ padding: 20, marginBottom: 14 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
        <Icon name="refresh" size={16} color="var(--cobalt)" />
        <span style={{ fontWeight: 700 }}>Parsing PCAP workbook…</span>
        <span style={{ marginLeft: "auto", fontSize: 11, color: "var(--ink-500)" }}>Mapping 142 columns</span>
      </div>
      <div className="progress-track"><div className="progress-fill" style={{ width: "70%" }}></div></div>
    </div>
  </div>
);

const HFErrorState = () => (
  <div style={{ padding: "30px 40px" }}>
    <div className="card" style={{ borderTop: "3px solid var(--purple)", padding: "24px 28px" }}>
      <div style={{ display: "flex", gap: 14 }}>
        <div style={{ background: "var(--purple)", color: "var(--white)", padding: 10 }}>
          <Icon name="warn" size={20} />
        </div>
        <div>
          <div style={{ fontSize: 10.5, fontWeight: 800, letterSpacing: "0.14em", textTransform: "uppercase", color: "var(--purple)" }}>
            Upload Rejected
          </div>
          <h2 style={{ margin: "4px 0 6px", fontSize: 18, color: "var(--kpmg-blue)" }}>
            PCAP header row not detected
          </h2>
          <p style={{ margin: 0, fontSize: 12.5, color: "var(--ink-500)", maxWidth: 540 }}>
            Row 1 of a PCAP workbook should be merged group headers; row 2 should contain canonical column names. Either row 2 is empty, or required columns are missing.
          </p>
        </div>
      </div>
    </div>
  </div>
);

window.HFTab = HFTab;
