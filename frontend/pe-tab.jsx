/* Private Equity tab */

const PETab = ({
  fileState, allInvestors, selectedInvestors, scope,
  generationState, setGenerationState,
  results, setResults,
  insightState, setInsightState,
  insightMd, setInsightMd,
  insightType, setInsightType,
  insightScope, setInsightScope,
  insightInvestor, setInsightInvestor,
  gridView, setGridView,
  darkNavyHero,
  onTriggerGenerate,
  onTriggerInsight,
  setMissingColsError, missingColsError,
  downloadUrls,
  runId,
  investorsData,
  partnership,
  asOf,
  generationError,
  insightError,
}) => {
  // ─── EARLY STATES
  if (fileState === "empty")   return <PEEmpty />;
  if (fileState === "error")   return <PEErrorState missingCols={missingColsError || ["INCEPTION_TO_DATE_CONTRIBUTION", "TEV_RATIO"]} />;
  if (fileState === "loading") return <PELoading />;

  const displayInvestors = investorsData || [];
  const partnershipName  = partnership || "—";
  const asOfLabel        = asOf        || "—";

  const targetInvestors = scope === "all"
    ? displayInvestors
    : displayInvestors.filter(i => selectedInvestors.includes(i.name));

  // ─── LOADED — metrics
  const totalCommit  = displayInvestors.reduce((s, i) => s + (i.committed   || 0), 0);
  const totalNav     = displayInvestors.reduce((s, i) => s + (i.closing_nav || 0), 0);
  const totalDist    = displayInvestors.reduce((s, i) => s + (i.itd_dist    || 0), 0);
  const totalContrib = displayInvestors.reduce((s, i) => s + (i.itd_contrib || 0), 0);
  const tvpi = totalContrib > 0 ? (totalDist + totalNav) / totalContrib : 0;
  const dpi  = totalContrib > 0 ? totalDist / totalContrib : 0;

  // Verdicts are "PENDING" until generate runs; treat PENDING as not-yet-validated
  const passes  = displayInvestors.filter(i => i.verdict === "ALL_PASS").length;
  const revals  = displayInvestors.filter(i => i.verdict === "REVALUABLE").length;
  const invalids = displayInvestors.filter(i => i.verdict === "INVALID").length;

  // Use real results' verdicts after generation has run
  const resultPasses  = results.filter(r => r.verdict === "ALL_PASS").length;
  const resultRevals  = results.filter(r => r.verdict === "REVALUABLE").length;
  const resultInvalids = results.filter(r => r.verdict === "INVALID").length;

  const dispPasses  = generationState === "done" ? resultPasses  : passes;
  const dispRevals  = generationState === "done" ? resultRevals  : revals;
  const dispInvalids = generationState === "done" ? resultInvalids : invalids;

  return (
    <div style={{ padding: "20px 24px 60px" }}>
      {/* Hero metric strip */}
      <div className={darkNavyHero ? "hero-darknav" : ""}
        style={{
          padding: darkNavyHero ? "20px 20px" : "0",
          marginBottom: 22,
          marginLeft: darkNavyHero ? -24 : 0,
          marginRight: darkNavyHero ? -24 : 0,
          marginTop: darkNavyHero ? -20 : 0,
          paddingLeft: darkNavyHero ? 24 : 0,
          paddingRight: darkNavyHero ? 24 : 0,
        }}>
        {darkNavyHero && (
          <div style={{ display: "flex", alignItems: "baseline", gap: 14, marginBottom: 14 }}>
            <h1 style={{ fontSize: 22, margin: 0, fontWeight: 800, letterSpacing: "-0.02em", color: "var(--white)" }}>
              {partnershipName}
            </h1>
            <span style={{ fontSize: 12, color: "var(--light-blue)" }}>Private Equity · As of {asOfLabel}</span>
          </div>
        )}
        {!darkNavyHero && (
          <div style={{ display: "flex", alignItems: "baseline", gap: 14, marginBottom: 14, flexWrap: "wrap" }}>
            <h1 style={{ fontSize: 20, margin: 0, fontWeight: 800, letterSpacing: "-0.02em", color: "var(--kpmg-blue)" }}>
              {partnershipName}
            </h1>
            <span style={{ fontSize: 11, color: "var(--ink-500)", letterSpacing: "0.06em", textTransform: "uppercase", fontWeight: 600 }}>
              Private Equity · {asOfLabel} · USD
            </span>
          </div>
        )}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, minmax(0, 1fr))", gap: 10 }}>
          <MetricCard label="Limited Partners"
            value={<span className="num">{displayInvestors.length}</span>}
            sub={<span>{scope === "all" ? "All investors" : `${targetInvestors.length} selected`}</span>}
            accent="pacific" />
          <MetricCard label="Committed"
            value={<span className="num">{window.fmtMoney(totalCommit)}</span>}
            sub={<span>{totalContrib > 0 ? `${((totalContrib/totalCommit)*100).toFixed(1)}% called` : "—"} · {window.fmtMoney(totalContrib)} contrib</span>}
            accent="cobalt" />
          <MetricCard label="Closing NAV"
            value={<span className="num">{window.fmtMoney(totalNav)}</span>}
            sub={<span>{window.fmtMoney(totalDist)} distributed ITD</span>} />
          <MetricCard label="TVPI · DPI"
            value={<span className="num">{tvpi.toFixed(2)}x · {dpi.toFixed(2)}x</span>}
            sub={<span><strong style={{color:"var(--teal)"}}>TEV</strong> {window.fmtMoney(totalDist + totalNav)}</span>}
            accent="teal" />
        </div>
      </div>

      {/* Selected pills + chips */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 18 }}>
        <span style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", color: "var(--ink-500)" }}>
          Selected ({targetInvestors.length})
        </span>
        {targetInvestors.slice(0, 6).map(inv => (
          <Pill key={inv.name} color="blue">{inv.name}</Pill>
        ))}
        {targetInvestors.length > 6 && <Pill>+{targetInvestors.length - 6} more</Pill>}
        <div style={{ flex: 1 }}></div>
        <button className="btn btn-ghost btn-sm">
          <Icon name="filter" size={12} /> Filters
        </button>
      </div>

      {/* Preview data grid */}
      <SecHeading right={
        <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
          <button className={`icon-btn`} onClick={() => setGridView("table")} style={gridView === "table" ? { background: "var(--kpmg-blue)", color: "var(--white)" } : {}} title="Table view">
            <Icon name="table" size={13} />
          </button>
          <button className={`icon-btn`} onClick={() => setGridView("cards")} style={gridView === "cards" ? { background: "var(--kpmg-blue)", color: "var(--white)" } : {}} title="Card view">
            <Icon name="grid" size={13} />
          </button>
        </div>
      }>Preview Source Data</SecHeading>

      <div className="card card-pacific" style={{ marginBottom: 22 }}>
        {gridView === "table"
          ? <PEDataGrid investors={displayInvestors} />
          : <PEDataCards investors={displayInvestors} />}
      </div>

      {/* GENERATE */}
      <SecHeading right={generationState === "done" && (
        <span style={{ fontSize: 10.5, color: "var(--teal)", fontWeight: 700, letterSpacing: "0.08em" }}>
          ● COMPLETE
        </span>
      )}>Generate Documents</SecHeading>

      {generationError && (
        <div style={{ marginBottom: 12, padding: "10px 14px", background: "#fdf2f4", border: "1px solid var(--purple)", display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: "var(--purple)" }}>
          <Icon name="warn" size={13} color="var(--purple)" />
          <span style={{ flex: 1 }}>{generationError}</span>
        </div>
      )}
      <div className="card" style={{ padding: 16, marginBottom: 22 }}>
        <PEGeneratePanel
          targetInvestors={targetInvestors}
          generationState={generationState}
          results={results}
          onTriggerGenerate={onTriggerGenerate}
          passes={dispPasses} revals={dispRevals} invalids={dispInvalids}
          runId={runId}
        />
      </div>

      {/* Downloads */}
      {generationState === "done" && (
        <>
          <SecHeading>Download Generated Documents</SecHeading>
          <div className="card" style={{ padding: 16, marginBottom: 22, background: "linear-gradient(180deg, rgba(0,184,245,0.04), transparent)" }}>
            <PEDownloads count={dispPasses + dispRevals} downloadUrls={downloadUrls} />
          </div>
        </>
      )}

      {/* AI Insights */}
      <SecHeading right={
        <span style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 10.5, color: "var(--ink-500)", fontWeight: 600 }}>
          <Icon name="spark" size={11} color="var(--cobalt)" />
          Gemini 2.5 Pro · 65K tokens
        </span>
      }>AI PE Insights</SecHeading>

      <div className="card" style={{ padding: 16 }}>
        {insightError && (
          <div style={{ marginBottom: 12, padding: "10px 14px", background: "#fdf2f4", border: "1px solid var(--purple)", display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: "var(--purple)" }}>
            <Icon name="warn" size={13} color="var(--purple)" />
            <span style={{ flex: 1 }}>{insightError}</span>
          </div>
        )}
        <PEInsights
          insightState={insightState}
          insightMd={insightMd}
          insightType={insightType} setInsightType={setInsightType}
          insightScope={insightScope} setInsightScope={setInsightScope}
          insightInvestor={insightInvestor} setInsightInvestor={setInsightInvestor}
          onTriggerInsight={onTriggerInsight}
          allInvestors={displayInvestors.map(i => i.name)}
        />
      </div>
    </div>
  );
};

const PEEmpty = () => (
  <div style={{ padding: "60px 40px" }}>
    <div className="card card-pacific" style={{ padding: "48px 40px", textAlign: "center", background: "linear-gradient(180deg, rgba(0,184,245,0.04), transparent 60%)" }}>
      <div style={{ display: "inline-flex", padding: 18, border: "1.5px dashed var(--border-strong)", marginBottom: 18 }}>
        <Icon name="upload" size={32} color="var(--kpmg-blue)" />
      </div>
      <h1 style={{ fontSize: 24, color: "var(--kpmg-blue)", margin: "0 0 8px", fontWeight: 800, letterSpacing: "-0.02em" }}>
        Upload an investor data file to begin
      </h1>
      <p style={{ color: "var(--ink-500)", fontSize: 13.5, maxWidth: 500, margin: "0 auto 18px", lineHeight: 1.55 }}>
        Drop a Private Equity capital data <code style={{ background: "var(--ink-100)", padding: "1px 6px" }}>.xlsx</code> in the sidebar.
        We'll wrangle types, validate arithmetic, and let you generate per-investor Word + PDF statements.
      </p>
      <div style={{ display: "inline-flex", gap: 8, marginTop: 6 }}>
        <button className="btn btn-primary" onClick={() => {
          const input = document.createElement("input");
          input.type = "file";
          input.accept = ".xlsx";
          input.onchange = (e) => {
            const file = e.target.files[0];
            if (file) window.onPeFileSelected(file);
          };
          input.click();
        }}>
          <Icon name="upload" size={13} /> Upload PE data
        </button>
        <button className="btn btn-ghost"><Icon name="file" size={13} /> Use sample data</button>
      </div>

      {/* Expected columns chip strip */}
      <div style={{ marginTop: 28 }}>
        <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", color: "var(--ink-500)", marginBottom: 8 }}>
          Expected columns
        </div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 4, justifyContent: "center", maxWidth: 720, margin: "0 auto" }}>
          {["INVESTOR_ID", "INVESTOR_NAME", "PARTNERSHIP_NAME", "CURRENCY_CODE", "FROM_DATE", "TO_DATE", "COMMITTED_CAPITAL",
            "INCEPTION_TO_DATE_CONTRIBUTION", "INCEPTION_TO_DATE_DISTRIBUTION", "OPENING_YTD_NAV",
            "CLOSING_YTD_NAV", "TEV", "TEV_RATIO", "MANAGEMENT_FEE", "INCENTIVE_FEE"]
            .map(c => (
              <span key={c} style={{
                fontSize: 10.5, fontFamily: "var(--font-mono)",
                padding: "2px 7px", border: "1px solid var(--border)",
                color: "var(--ink-700)", background: "var(--white)"
              }}>{c}</span>
            ))}
        </div>
      </div>

      {/* Pipeline diagram */}
      <div style={{ marginTop: 32, display: "flex", alignItems: "center", justifyContent: "center", gap: 14, flexWrap: "wrap" }}>
        {[
          { i: "upload",   t: "Upload",    s: "xlsx" },
          { i: "refresh",  t: "Wrangle",   s: "type coercion" },
          { i: "shield",   t: "Validate",  s: "4 checks + Gemini" },
          { i: "doc",      t: "Generate",  s: "Word · PDF" },
          { i: "download", t: "Package",   s: "ZIP + Summary" },
        ].map((step, idx, arr) => (
          <React.Fragment key={step.t}>
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", minWidth: 80 }}>
              <div style={{ width: 36, height: 36, background: "var(--kpmg-blue)", color: "var(--white)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Icon name={step.i} size={16} />
              </div>
              <div style={{ marginTop: 6, fontSize: 11.5, fontWeight: 700, color: "var(--ink-900)" }}>{step.t}</div>
              <div style={{ fontSize: 10, color: "var(--ink-500)" }}>{step.s}</div>
            </div>
            {idx < arr.length - 1 && <Icon name="chevron-right" size={12} color="var(--ink-300)" />}
          </React.Fragment>
        ))}
      </div>
    </div>
  </div>
);

const PEErrorState = ({ missingCols }) => (
  <div style={{ padding: "30px 40px" }}>
    <div className="card" style={{ borderTop: "3px solid var(--purple)", padding: "24px 28px" }}>
      <div style={{ display: "flex", alignItems: "flex-start", gap: 14 }}>
        <div style={{ background: "var(--purple)", color: "var(--white)", padding: 10, display: "inline-flex" }}>
          <Icon name="warn" size={20} />
        </div>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 10.5, fontWeight: 800, letterSpacing: "0.14em", textTransform: "uppercase", color: "var(--purple)" }}>
            Upload Rejected
          </div>
          <h2 style={{ margin: "4px 0 6px", fontSize: 18, color: "var(--kpmg-blue)" }}>
            Required columns missing from uploaded file
          </h2>
          <p style={{ margin: 0, fontSize: 12.5, color: "var(--ink-500)" }}>
            The PE wrangler expects all canonical column headers (row 1). Add these columns and re-upload:
          </p>
          <div style={{ marginTop: 14, display: "flex", flexWrap: "wrap", gap: 6 }}>
            {missingCols.map(c => (
              <span key={c} style={{
                fontFamily: "var(--font-mono)", fontSize: 11,
                background: "rgba(114,19,234,0.08)", color: "var(--purple)",
                padding: "3px 8px", border: "1px solid rgba(114,19,234,0.25)"
              }}>
                <Icon name="x" size={10} /> &nbsp;{c}
              </span>
            ))}
          </div>
          <div style={{ marginTop: 18, display: "flex", gap: 8 }}>
            <button className="btn btn-primary"><Icon name="upload" size={12} /> Re-upload file</button>
            <button className="btn btn-ghost"><Icon name="external" size={12} /> View column spec</button>
          </div>
        </div>
      </div>
    </div>
  </div>
);

const PELoading = () => (
  <div style={{ padding: 30 }}>
    <div style={{ display: "grid", gridTemplateColumns: "repeat(6,1fr)", gap: 10, marginBottom: 22 }}>
      {Array.from({ length: 6 }).map((_, i) => (
        <div key={i} className="metric">
          <div className="skeleton" style={{ height: 9, width: "60%" }}></div>
          <div className="skeleton" style={{ height: 22, width: "90%" }}></div>
          <div className="skeleton" style={{ height: 8, width: "40%" }}></div>
        </div>
      ))}
    </div>
    <div className="card" style={{ padding: 20, marginBottom: 14 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
        <Icon name="refresh" size={16} color="var(--cobalt)" />
        <span style={{ fontWeight: 700 }}>Parsing investor data…</span>
        <span style={{ marginLeft: "auto", fontSize: 11, color: "var(--ink-500)" }}>Stage 2 of 3</span>
      </div>
      <div className="progress-track"><div className="progress-fill" style={{ width: "62%" }}></div></div>
      <div style={{ marginTop: 14, display: "grid", gridTemplateColumns: "auto 1fr auto", gap: "6px 12px", fontSize: 11.5, alignItems: "center" }}>
        {[
          ["check",   "Read xlsx — wrangling types, filling NaN",    "var(--teal)",   "complete"],
          ["refresh", "Grouping by INVESTOR_NAME · last period",      "var(--cobalt)", "in progress"],
          ["dot",     "Building investor selection cache",            "var(--ink-300)","queued"],
        ].map(([i, t, c, s], idx) => (
          <React.Fragment key={idx}>
            <Icon name={i} size={13} color={c} stroke={2.2} />
            <span style={{ color: "var(--ink-700)" }}>{t}</span>
            <span style={{ color: "var(--ink-500)", fontFamily: "var(--font-mono)", fontSize: 10.5, textTransform: "uppercase" }}>{s}</span>
          </React.Fragment>
        ))}
      </div>
    </div>
  </div>
);

const PEDataGrid = ({ investors = [] }) => (
  <div style={{ maxHeight: 360, overflow: "auto" }}>
    <table className="dgrid">
      <thead>
        <tr>
          <th>Investor</th>
          <th>ID</th>
          <th>CCY</th>
          <th className="right">Committed</th>
          <th className="right">ITD Contrib</th>
          <th className="right">ITD Dist</th>
          <th className="right">Opening NAV</th>
          <th className="right">Closing NAV</th>
          <th className="right">TEV</th>
          <th className="right">TEV ratio</th>
          <th>Verdict</th>
        </tr>
      </thead>
      <tbody>
        {investors.map((inv, i) => (
          <tr key={inv.investor_id || inv.name} className={i === 0 ? "selected" : ""}>
            <td><span style={{ fontWeight: 600 }}>{inv.name}</span></td>
            <td style={{ fontFamily: "var(--font-mono)", fontSize: 11, color: "var(--ink-500)" }}>{inv.investor_id || "—"}</td>
            <td>{inv.ccy || "USD"}</td>
            <td className="right">{window.fmtMoneyFull(inv.committed   || 0, inv.ccy)}</td>
            <td className="right">{window.fmtMoneyFull(inv.itd_contrib  || 0, inv.ccy)}</td>
            <td className="right">{window.fmtMoneyFull(inv.itd_dist     || 0, inv.ccy)}</td>
            <td className="right">{window.fmtMoneyFull(inv.opening_nav  || 0, inv.ccy)}</td>
            <td className="right">{window.fmtMoneyFull(inv.closing_nav  || 0, inv.ccy)}</td>
            <td className="right">{window.fmtMoneyFull(inv.tev          || 0, inv.ccy)}</td>
            <td className="right">{(inv.tev_ratio || 0).toFixed(3)}</td>
            <td>
              {inv.verdict && inv.verdict !== "PENDING"
                ? <Badge verdict={inv.verdict} size="sm" />
                : <span style={{ fontSize: 10, color: "var(--ink-400)", fontFamily: "var(--font-mono)" }}>—</span>
              }
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  </div>
);

const PEDataCards = ({ investors = [] }) => (
  <div style={{ maxHeight: 360, overflow: "auto", padding: 12, display: "grid", gridTemplateColumns: "repeat(2,1fr)", gap: 8 }}>
    {investors.slice(0, 12).map(inv => (
      <div key={inv.investor_id || inv.name} style={{
        background: "var(--white)", border: "1px solid var(--border)",
        padding: 12, display: "flex", flexDirection: "column", gap: 8
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <div>
            <div style={{ fontWeight: 700, fontSize: 13 }}>{inv.name}</div>
            <div style={{ fontSize: 10.5, color: "var(--ink-500)", fontFamily: "var(--font-mono)" }}>{inv.investor_id || "—"} · {inv.ccy || "USD"}</div>
          </div>
          <div style={{ marginLeft: "auto" }}>
            {inv.verdict && inv.verdict !== "PENDING"
              ? <Badge verdict={inv.verdict} size="sm" />
              : <span style={{ fontSize: 10, color: "var(--ink-400)" }}>pending</span>
            }
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 6, fontSize: 11 }}>
          {[
            ["COMMIT",    window.fmtMoney(inv.committed   || 0, inv.ccy)],
            ["NAV",       window.fmtMoney(inv.closing_nav || 0, inv.ccy)],
            ["TEV ratio", (inv.tev_ratio || 0).toFixed(3)],
          ].map(([k, v]) => (
            <div key={k}>
              <div style={{ fontSize: 9.5, letterSpacing: "0.1em", color: "var(--ink-500)", textTransform: "uppercase" }}>{k}</div>
              <div className="num" style={{ fontWeight: 700, color: "var(--kpmg-blue)" }}>{v}</div>
            </div>
          ))}
        </div>
      </div>
    ))}
  </div>
);

const PEGeneratePanel = ({ targetInvestors, generationState, results, onTriggerGenerate, passes, revals, invalids, runId }) => {
  if (generationState === "idle") {
    return (
      <div>
        <div style={{ display: "flex", alignItems: "flex-start", gap: 18, marginBottom: 14 }}>
          <div style={{ flex: 1 }}>
            <h3 style={{ margin: "0 0 6px", fontSize: 16, color: "var(--ink-900)" }}>
              Generate {targetInvestors.length} capital statement{targetInvestors.length === 1 ? "" : "s"}
            </h3>
            <p style={{ margin: 0, fontSize: 12.5, color: "var(--ink-500)", maxWidth: 560 }}>
              Runs wrangler → validate → generate. Each investor produces a Word doc and PDF. INVALID rows are skipped and logged to the audit trail.
            </p>
            <div style={{ marginTop: 12, display: "flex", flexWrap: "wrap", gap: 16, alignItems: "center" }}>
              <ValidationPreview passes={passes} revals={revals} invalids={invalids} />
            </div>
          </div>
          <button className="btn btn-primary btn-lg" onClick={onTriggerGenerate}>
            <Icon name="play" size={14} />
            Generate {targetInvestors.length} statement{targetInvestors.length === 1 ? "" : "s"}
          </button>
        </div>

        <div style={{ display: "flex", gap: 16, fontSize: 11, color: "var(--ink-500)", paddingTop: 12, borderTop: "1px solid var(--border)" }}>
          <span><Icon name="shield" size={11} color="var(--teal)" /> &nbsp;4 arithmetic checks per row</span>
          <span><Icon name="spark" size={11} color="var(--cobalt)" /> &nbsp;Gemini 2.5 Pro 2nd opinion</span>
          <span><Icon name="audit" size={11} color="var(--ink-700)" /> &nbsp;JSONL audit log</span>
          <span style={{ marginLeft: "auto" }}>Est. <strong style={{ color: "var(--ink-900)" }}>~{Math.round(targetInvestors.length * 1.8)}s</strong></span>
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
          <span style={{ fontWeight: 700, fontSize: 14 }}>Generating documents…</span>
          <span style={{ marginLeft: "auto", fontSize: 12, color: "var(--ink-500)", fontFamily: "var(--font-mono)" }}>
            {results?.length || 0} / {targetInvestors.length}
          </span>
        </div>
        <div className="progress-track"><div className="progress-fill" style={{ width: `${progress * 100}%` }}></div></div>
        <div style={{ marginTop: 14, maxHeight: 240, overflow: "auto" }}>
          {(results || []).slice().reverse().map((r, i) => <ResultLine key={i} r={r} />)}
          {results?.length < targetInvestors.length && (
            <div className="result-row" style={{ opacity: 0.7 }}>
              <Icon name="refresh" size={14} color="var(--cobalt)" stroke={2.4} />
              <div>
                <div className="r-name pulse">Processing {targetInvestors[results?.length || 0]?.name}…</div>
                <div className="r-file">validating arithmetic · waiting for Gemini</div>
              </div>
              <span></span>
              <span></span>
            </div>
          )}
        </div>
      </div>
    );
  }

  // done
  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 14 }}>
        <div style={{ background: "var(--teal)", color: "var(--white)", padding: 10, display: "inline-flex" }}>
          <Icon name="check" size={18} />
        </div>
        <div>
          <div style={{ fontSize: 14, fontWeight: 700, color: "var(--ink-900)" }}>
            Generated {results.filter(r => r.ok).length} of {results.length} statements
          </div>
          <div style={{ fontSize: 12, color: "var(--ink-500)" }}>
            Run · <span style={{ fontFamily: "var(--font-mono)" }}>{runId || "a4f1-c290"}</span> · Audit log appended
          </div>
        </div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 6 }}>
          <button className="btn btn-ghost btn-sm"><Icon name="audit" size={12} /> View audit log</button>
          <button className="btn btn-ghost btn-sm"><Icon name="refresh" size={12} /> Re-run</button>
        </div>
      </div>
      <div style={{ maxHeight: 280, overflow: "auto", border: "1px solid var(--border)" }}>
        {results.map((r, i) => <ResultLine key={i} r={r} />)}
      </div>
    </div>
  );
};

const ValidationPreview = ({ passes, revals, invalids }) => {
  const total = passes + revals + invalids;
  return (
    <div style={{ display: "flex", alignItems: "stretch", gap: 8, fontVariantNumeric: "tabular-nums" }}>
      <div style={{ display: "flex", flexDirection: "column", justifyContent: "center" }}>
        <div style={{ fontSize: 9.5, letterSpacing: "0.12em", color: "var(--ink-500)", fontWeight: 700, textTransform: "uppercase" }}>
          Pre-flight Validation
        </div>
        <div style={{ display: "flex", height: 8, marginTop: 6, width: 220 }}>
          <div style={{ flex: passes  || 0.01, background: "var(--teal)" }}></div>
          <div style={{ flex: revals  || 0,    background: "var(--pacific)" }}></div>
          <div style={{ flex: invalids || 0,   background: "var(--purple)" }}></div>
        </div>
      </div>
      <div style={{ display: "flex", gap: 6 }}>
        <ValidationTile verdict="ALL_PASS"   count={passes}   total={total} />
        <ValidationTile verdict="REVALUABLE" count={revals}   total={total} />
        <ValidationTile verdict="INVALID"    count={invalids} total={total} />
      </div>
    </div>
  );
};

const ResultLine = ({ r }) => {
  const apiBase = window.location.port === "8080" ? "http://localhost:8000" : "";
  return (
    <div className="result-row">
      {r.ok ? (
        <Icon name="check" size={14} color="var(--teal)" stroke={2.5} />
      ) : (
        <Icon name="x" size={14} color="var(--purple)" stroke={2.5} />
      )}
      <div style={{ minWidth: 0 }}>
        <div className="r-name">{r.investor}</div>
        <div className="r-file">{r.file || r.error}</div>
      </div>
      <Badge verdict={r.verdict} size="sm" />
      <div style={{ display: "flex", gap: 4 }}>
        {r.ok && <>
          {r.doc_url && (
            <a href={`${apiBase}${r.doc_url}`} download>
              <button className="icon-btn" title="Download Word"><Icon name="doc" size={11} /></button>
            </a>
          )}
          {r.pdf_url && (
            <a href={`${apiBase}${r.pdf_url}`} download>
              <button className="icon-btn" title="Download PDF"><Icon name="file" size={11} /></button>
            </a>
          )}
        </>}
      </div>
    </div>
  );
};

const PEDownloads = ({ count, downloadUrls }) => (
  <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10 }}>
    {[
      { i: "doc",      t: "Word ZIP",      s: `${count} .docx files`,      cls: "btn-accent", urlKey: "word_zip" },
      { i: "file",     t: "PDF ZIP",       s: `${count} .pdf files`,       cls: "btn-accent", urlKey: "pdf_zip" },
      { i: "table",    t: "Summary Excel", s: "All investors · validated", cls: "btn-ghost",  urlKey: "summary_excel" },
      { i: "download", t: "Summary CSV",   s: "Pandas export",             cls: "btn-ghost",  urlKey: "summary_csv" },
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

const PEInsights = ({
  insightState, insightMd, insightType, setInsightType,
  insightScope, setInsightScope, insightInvestor, setInsightInvestor,
  onTriggerInsight, allInvestors,
}) => (
  <div>
    {/* Controls */}
    <div style={{ display: "grid", gridTemplateColumns: "1fr 1.4fr", gap: 16, marginBottom: 16 }}>
      <div>
        <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: "var(--ink-500)", marginBottom: 6 }}>Scope</div>
        <div style={{ display: "flex", border: "1px solid var(--border-strong)" }}>
          <button onClick={() => setInsightScope("portfolio")}
            style={{
              flex: 1, padding: "10px", border: "none", background: insightScope === "portfolio" ? "var(--kpmg-blue)" : "var(--white)",
              color: insightScope === "portfolio" ? "var(--white)" : "var(--ink-700)",
              fontWeight: 600, fontSize: 12, cursor: "pointer"
            }}>
            <Icon name="users" size={12} /> &nbsp;Portfolio Overview
          </button>
          <button onClick={() => setInsightScope("investor")}
            style={{
              flex: 1, padding: "10px", border: "none", background: insightScope === "investor" ? "var(--kpmg-blue)" : "var(--white)",
              color: insightScope === "investor" ? "var(--white)" : "var(--ink-700)",
              fontWeight: 600, fontSize: 12, cursor: "pointer"
            }}>
            <Icon name="user" size={12} /> &nbsp;Individual Investor
          </button>
        </div>
        {insightScope === "investor" && (
          <select value={insightInvestor} onChange={e => setInsightInvestor(e.target.value)}
            style={{ width: "100%", marginTop: 8 }}>
            {(allInvestors || []).map(name => (
              <option key={name}>{name}</option>
            ))}
          </select>
        )}
      </div>
      <div>
        <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: "var(--ink-500)", marginBottom: 6 }}>Insight Type</div>
        <div className="chip-select">
          {window.PE_INSIGHT_TYPES.map(t => (
            <button
              key={t.key}
              className={`chip ${insightType === t.key ? "active" : ""}`}
              onClick={() => setInsightType(t.key)}>
              {t.label}
            </button>
          ))}
        </div>
      </div>
    </div>

      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
        <button className="btn btn-dark btn-lg" onClick={onTriggerInsight} disabled={insightState === "running"}>
          <Icon name="spark" size={14} />
          {insightState === "running" ? "Analyzing…" : "Generate PE Insights"}
        </button>
        <span style={{ fontSize: 11, color: "var(--ink-500)" }}>
          Senior PE advisor persona
        </span>
        {insightState === "done" && (
          <div style={{ marginLeft: "auto", display: "flex", gap: 4 }}>
            <button className="btn btn-ghost btn-sm"><Icon name="copy" size={11} /> Copy</button>
            <button className="btn btn-ghost btn-sm"><Icon name="download" size={11} /> Export MD</button>
          </div>
        )}
      </div>

    {/* Output */}
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
          <span>Persona: Senior PE Advisor</span>
          <span style={{ marginLeft: "auto" }}>Logged to audit trail · <span style={{ fontFamily: "var(--font-mono)" }}>gemini_insight</span></span>
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

const InsightLoading = () => (
  <div style={{ padding: "16px 20px", border: "1px solid var(--border)", borderTop: "3px solid var(--cobalt)", background: "rgba(30,73,226,0.03)" }}>
    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
      <Icon name="spark" size={14} color="var(--cobalt)" stroke={2.2} />
      <span className="pulse" style={{ fontSize: 12.5, fontWeight: 600 }}>Gemini 2.5 Pro is analyzing the portfolio…</span>
      <span style={{ marginLeft: "auto", fontSize: 11, color: "var(--ink-500)", fontFamily: "var(--font-mono)" }}>~3-5s</span>
    </div>
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      <div className="skeleton" style={{ height: 14, width: "30%" }}></div>
      <div className="skeleton" style={{ height: 10, width: "92%" }}></div>
      <div className="skeleton" style={{ height: 10, width: "88%" }}></div>
      <div className="skeleton" style={{ height: 10, width: "75%" }}></div>
      <div className="skeleton" style={{ height: 14, width: "26%", marginTop: 8 }}></div>
      <div className="skeleton" style={{ height: 10, width: "94%" }}></div>
      <div className="skeleton" style={{ height: 10, width: "70%" }}></div>
    </div>
  </div>
);

window.PETab = PETab;
window.PEEmpty = PEEmpty;
