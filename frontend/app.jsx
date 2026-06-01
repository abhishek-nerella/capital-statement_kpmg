/* App root — wires state for both PE + HF, integrates FastAPI backend */

const API_BASE = window.location.port === "8080" ? "http://localhost:8000" : "";

const TWEAK_DEFAULTS = /*EDITMODE-BEGIN*/{
  "theme": "light",
  "denseGrid": true
}/*EDITMODE-END*/;

// ── Backend error banner ───────────────────────────────────────────────────────
const BackendErrorBanner = ({ message, onDismiss }) => (
  <div style={{
    display: "flex", alignItems: "center", gap: 10,
    padding: "10px 20px", background: "#fdf2f4",
    borderBottom: "2px solid var(--purple)",
    fontSize: 12.5, color: "var(--purple)", fontWeight: 600,
  }}>
    <Icon name="warn" size={14} color="var(--purple)" />
    <span style={{ flex: 1 }}>{message}</span>
    <button onClick={onDismiss} style={{
      background: "none", border: "none", cursor: "pointer",
      color: "var(--purple)", display: "flex", padding: 2,
    }}>
      <Icon name="x" size={13} />
    </button>
  </div>
);

const App = () => {
  const { isAuthenticated, login, logout } = useAuth();

  // ── Gate: show login until authenticated ──────────────────────────────────
  if (!isAuthenticated) {
    return <LoginPage onLogin={login} />;
  }

  const [t, setTweak] = useTweaks(TWEAK_DEFAULTS);
  const darkNavyHero = t.theme === "dark";

  // ── PE state ──────────────────────────────────────────────────────────────
  const [peFileState,       setPeFileState]       = React.useState("empty");
  const [peScope,           setPeScope]           = React.useState("all");
  const [peSelected,        setPeSelected]        = React.useState([]);
  const [peGenerationState, setPeGenerationState] = React.useState("idle");
  const [peGenerationError, setPeGenerationError] = React.useState(null);
  const [peResults,         setPeResults]         = React.useState([]);
  const [peGridView,        setPeGridView]        = React.useState("table");
  const [peInsightState,    setPeInsightState]    = React.useState("idle");
  const [peInsightError,    setPeInsightError]    = React.useState(null);
  const [peInsightMd,       setPeInsightMd]       = React.useState("");
  const [peInsightType,     setPeInsightType]     = React.useState("full");
  const [peInsightScope,    setPeInsightScope]    = React.useState("portfolio");
  const [peInsightInvestor, setPeInsightInvestor] = React.useState("");

  // PE API state
  const [peSessionToken,  setPeSessionToken]  = React.useState(null);
  const [peInvestorList,  setPeInvestorList]  = React.useState([]);
  const [pePreviewData,   setPePreviewData]   = React.useState(null);
  const [pePartnership,   setPePartnership]   = React.useState(null);
  const [peAsOf,          setPeAsOf]          = React.useState(null);
  const [peFileName,      setPeFileName]      = React.useState(null);
  const [peDownloadUrls,  setPeDownloadUrls]  = React.useState(null);
  const [peRunId,         setPeRunId]         = React.useState(null);
  const [peMissingCols,   setPeMissingCols]   = React.useState([]);
  const [peSessionData,   setPeSessionData]   = React.useState(null);

  // ── HF state ──────────────────────────────────────────────────────────────
  const [hfFileState,       setHfFileState]       = React.useState("empty");
  const [hfScope,           setHfScope]           = React.useState("all");
  const [hfSelected,        setHfSelected]        = React.useState([]);
  const [hfGenerationState, setHfGenerationState] = React.useState("idle");
  const [hfGenerationError, setHfGenerationError] = React.useState(null);
  const [hfResults,         setHfResults]         = React.useState([]);
  const [hfHasLedger,       setHfHasLedger]       = React.useState(false);
  const [hfInsightState,    setHfInsightState]    = React.useState("idle");
  const [hfInsightError,    setHfInsightError]    = React.useState(null);
  const [hfInsightMd,       setHfInsightMd]       = React.useState("");
  const [hfInsightType,     setHfInsightType]     = React.useState("full");
  const [hfInsightScope,    setHfInsightScope]    = React.useState("portfolio");
  const [hfInsightInvestor, setHfInsightInvestor] = React.useState("");

  // HF API state
  const [hfPcapToken,      setHfPcapToken]      = React.useState(null);
  const [hfInvestorList,   setHfInvestorList]   = React.useState([]);
  const [hfPreviewData,    setHfPreviewData]     = React.useState(null);
  const [hfFundStats,      setHfFundStats]       = React.useState(null);
  const [hfLedgerToken,    setHfLedgerToken]     = React.useState(null);
  const [hfFileName,       setHfFileName]        = React.useState(null);
  const [hfLedgerFileName, setHfLedgerFileName]  = React.useState(null);
  const [hfDownloadUrls,   setHfDownloadUrls]    = React.useState(null);
  const [hfRunId,          setHfRunId]           = React.useState(null);
  const [hfSessionData,    setHfSessionData]      = React.useState(null);

  // Global connection error banner
  const [connectionError, setConnectionError] = React.useState(null);

  React.useEffect(() => {
    window.onHfFileSelected = handleHfFileSelected;
    window.onHfLedgerSelected = handleHfLedgerSelected;
    window.onPeFileSelected = handlePeFileSelected;
  }, [peSessionToken, hfPcapToken]); // re-bind if tokens change if needed, though functions are stable

  // Tabs + drawers
  const [activeTab,  setActiveTab]  = React.useState("pe");
  const [chatOpen,   setChatOpen]   = React.useState(false);
  const [auditOpen,  setAuditOpen]  = React.useState(false);

  // ── PE handlers ───────────────────────────────────────────────────────────

  const handlePeFileSelected = async (file) => {
    setPeFileState("loading");
    setPeFileName(file.name);
    setConnectionError(null);
    try {
      const fd = new FormData();
      fd.append("file", file);
      const res = await fetch(`${API_BASE}/api/pe/upload`, { method: "POST", body: fd });
      const data = await res.json();
      if (!data.ok) {
        setPeMissingCols(data.missing_cols || []);
        setPeFileState("error");
        return;
      }
      setPeSessionToken(data.session_token);
      setPeInvestorList(data.investors || []);
      setPeSelected(data.investors || []);
      setPePreviewData(data.preview || null);
      setPePartnership(data.partnership || null);
      setPeAsOf(data.as_of || null);
      setPeSessionData(data.session_data || null);
      setPeFileState("loaded");
    } catch (err) {
      setConnectionError(`PE upload failed — backend unreachable at ${API_BASE}. Is the FastAPI server running? (${err.message})`);
      setPeFileState("error");
    }
  };

  const handlePeClear = () => {
    setPeFileState("empty");
    setPeGenerationState("idle");
    setPeGenerationError(null);
    setPeResults([]);
    setPeInsightState("idle");
    setPeInsightError(null);
    setPeInsightMd("");
    setPeSessionToken(null);
    setPeInvestorList([]);
    setPeSelected([]);
    setPePreviewData(null);
    setPePartnership(null);
    setPeAsOf(null);
    setPeFileName(null);
    setPeDownloadUrls(null);
    setPeRunId(null);
    setPeSessionData(null);
    setConnectionError(null);
  };

  const handlePeGenerate = async () => {
    if (!peSessionToken) {
      setPeGenerationError("No file uploaded — please upload a PE data file first.");
      return;
    }
    setPeGenerationState("running");
    setPeGenerationError(null);
    setPeResults([]);
    setConnectionError(null);

    try {
      const investors = peScope === "all" ? peInvestorList : peSelected.filter(n => peInvestorList.includes(n));
      const res = await fetch(`${API_BASE}/api/pe/generate`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ session_token: peSessionToken, investors, session_data: peSessionData }),
      });

      const reader = res.body.getReader();
      const dec = new TextDecoder();
      let buf = "";
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buf += dec.decode(value, { stream: true });
        const lines = buf.split("\n");
        buf = lines.pop();
        for (const line of lines) {
          if (!line.startsWith("data: ")) continue;
          const evt = JSON.parse(line.slice(6));
          if (evt.type === "progress") {
            setPeResults(prev => [...prev, {
              investor: evt.investor, ok: evt.ok,
              verdict: evt.verdict, file: evt.file, error: evt.error,
              doc_url: evt.doc_url, pdf_url: evt.pdf_url,
            }]);
          } else if (evt.type === "done") {
            setPeRunId(evt.run_id);
            setPeDownloadUrls({
              word_zip:      `${API_BASE}${evt.word_zip_url}`,
              pdf_zip:       `${API_BASE}${evt.pdf_zip_url}`,
              summary_excel: `${API_BASE}${evt.summary_excel_url}`,
              summary_csv:   `${API_BASE}${evt.summary_csv_url}`,
            });
            setPeGenerationState("done");
          }
        }
      }
    } catch (err) {
      setPeGenerationError(`Generation failed — backend unreachable. ${err.message}`);
      setPeGenerationState("error");
      setConnectionError(`Backend unreachable at ${API_BASE} — ensure the FastAPI server is running on port 8000.`);
    }
  };

  const handlePeInsight = async () => {
    if (!peSessionToken) {
      setPeInsightError("No file uploaded — please upload a PE data file first.");
      return;
    }
    setPeInsightState("running");
    setPeInsightError(null);
    setConnectionError(null);

    try {
      const res = await fetch(`${API_BASE}/api/pe/insights`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          session_token: peSessionToken,
          scope:         peInsightScope,
          investor:      peInsightScope === "investor" ? peInsightInvestor : "",
          insight_type:  peInsightType,
          session_data:  peSessionData,
        }),
      });
      const data = await res.json();
      if (data.ok) {
        setPeInsightMd(data.markdown);
        setPeInsightState("done");
      } else {
        setPeInsightError(data.error || "Insight generation failed.");
        setPeInsightState("error");
      }
    } catch (err) {
      setPeInsightError(`Backend unreachable — ${err.message}`);
      setPeInsightState("error");
      setConnectionError(`Backend unreachable at ${API_BASE} — ensure the FastAPI server is running on port 8000.`);
    }
  };

  // ── HF handlers ───────────────────────────────────────────────────────────

  const handleHfFileSelected = async (file) => {
    setHfFileState("loading");
    setHfFileName(file.name);
    setConnectionError(null);
    try {
      const fd = new FormData();
      fd.append("file", file);
      const res = await fetch(`${API_BASE}/api/hf/upload-pcap`, { method: "POST", body: fd });
      const data = await res.json();
      if (!data.ok) {
        setHfFileState("error");
        return;
      }
      setHfPcapToken(data.pcap_token);
      setHfInvestorList(data.investors || []);
      setHfSelected(data.investors || []);
      setHfPreviewData(data.preview || null);
      setHfFundStats(data.n_investors ? {
        n_investors:  data.n_investors,
        total_nav_cq: data.total_nav_cq,
        avg_net_irr:  data.avg_net_irr,
        avg_tvpi:     data.avg_tvpi,
      } : null);
      setHfSessionData(data.session_data || null);
      setHfFileState("loaded");
    } catch (err) {
      setConnectionError(`HF upload failed — backend unreachable at ${API_BASE}. Is the FastAPI server running? (${err.message})`);
      setHfFileState("error");
    }
  };

  const handleHfLedgerSelected = async (file) => {
    setHfLedgerFileName(file.name);
    setConnectionError(null);
    try {
      const fd = new FormData();
      fd.append("file", file);
      const res = await fetch(`${API_BASE}/api/hf/upload-ledger`, { method: "POST", body: fd });
      const data = await res.json();
      if (data.ok) {
        setHfLedgerToken(data.ledger_token);
        setHfHasLedger(true);
      } else {
        setHfLedgerFileName(null);
      }
    } catch (err) {
      setConnectionError(`Ledger upload failed — backend unreachable. (${err.message})`);
      setHfLedgerFileName(null);
    }
  };

  const handleHfGenerate = async () => {
    if (!hfPcapToken) {
      setHfGenerationError("No PCAP file uploaded — please upload a PCAP file first.");
      return;
    }
    setHfGenerationState("running");
    setHfGenerationError(null);
    setHfResults([]);
    setConnectionError(null);

    try {
      const investors = hfScope === "all" ? hfInvestorList : hfSelected.filter(n => hfInvestorList.includes(n));
      const res = await fetch(`${API_BASE}/api/hf/generate`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pcap_token: hfPcapToken, ledger_token: hfLedgerToken, investors, session_data: hfSessionData }),
      });

      const reader = res.body.getReader();
      const dec = new TextDecoder();
      let buf = "";
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buf += dec.decode(value, { stream: true });
        const lines = buf.split("\n");
        buf = lines.pop();
        for (const line of lines) {
          if (!line.startsWith("data: ")) continue;
          const evt = JSON.parse(line.slice(6));
          if (evt.type === "progress") {
            setHfResults(prev => [...prev, {
              investor: evt.investor, ok: evt.ok,
              verdict: evt.verdict || "ALL_PASS", file: evt.file, error: evt.error,
              doc_url: evt.doc_url, pdf_url: evt.pdf_url,
            }]);
          } else if (evt.type === "done") {
            setHfRunId(evt.run_id);
            setHfDownloadUrls({
              word_zip:       `${API_BASE}${evt.word_zip_url}`,
              pdf_zip:        `${API_BASE}${evt.pdf_zip_url}`,
              summary_excel:  `${API_BASE}${evt.summary_excel_url}`,
              companion_xlsx: evt.companion_xlsx_url ? `${API_BASE}${evt.companion_xlsx_url}` : null,
            });
            setHfGenerationState("done");
          }
        }
      }
    } catch (err) {
      setHfGenerationError(`Generation failed — backend unreachable. ${err.message}`);
      setHfGenerationState("error");
      setConnectionError(`Backend unreachable at ${API_BASE} — ensure the FastAPI server is running on port 8000.`);
    }
  };

  const handleHfInsight = async () => {
    if (!hfPcapToken) {
      setHfInsightError("No PCAP file uploaded — please upload a PCAP file first.");
      return;
    }
    setHfInsightState("running");
    setHfInsightError(null);
    setConnectionError(null);

    try {
      const res = await fetch(`${API_BASE}/api/hf/insights`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          pcap_token:   hfPcapToken,
          scope:        hfInsightScope,
          investor:     hfInsightScope === "investor" ? hfInsightInvestor : "",
          insight_type: hfInsightType,
          session_data: hfSessionData,
        }),
      });
      const data = await res.json();
      if (data.ok) {
        setHfInsightMd(data.markdown);
        setHfInsightState("done");
      } else {
        setHfInsightError(data.error || "Insight generation failed.");
        setHfInsightState("error");
      }
    } catch (err) {
      setHfInsightError(`Backend unreachable — ${err.message}`);
      setHfInsightState("error");
      setConnectionError(`Backend unreachable at ${API_BASE} — ensure the FastAPI server is running on port 8000.`);
    }
  };

  // ── Sidebar wiring ────────────────────────────────────────────────────────
  const isPe = activeTab === "pe";

  const sidebarProps = isPe ? {
    fileState:            peFileState,
    fileName:             peFileName,
    onFileSelected:       handlePeFileSelected,
    onClearFile:          handlePeClear,
    scope:                peScope,
    setScope:             setPeScope,
    allInvestors:         peInvestorList,
    selectedInvestors:    peSelected,
    setSelectedInvestors: setPeSelected,
    onOpenChat:           () => setChatOpen(true),
    activeTab,
  } : {
    fileState:            hfFileState,
    fileName:             hfFileName,
    onFileSelected:       handleHfFileSelected,
    onClearFile:          () => {
      setHfFileState("empty");
      setHfPcapToken(null);
      setHfInvestorList([]);
      setHfSelected([]);
      setHfPreviewData(null);
      setHfFundStats(null);
      setHfLedgerToken(null);
      setHfHasLedger(false);
      setHfFileName(null);
      setHfLedgerFileName(null);
      setHfGenerationState("idle");
      setHfGenerationError(null);
      setHfResults([]);
      setHfSessionData(null);
      setConnectionError(null);
    },
    scope:                hfScope,
    setScope:             setHfScope,
    allInvestors:         hfInvestorList,
    selectedInvestors:    hfSelected,
    setSelectedInvestors: setHfSelected,
    onOpenChat:           () => setChatOpen(true),
    activeTab,
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", overflow: "hidden", position: "relative" }}>
      <TopHeader
        runId={peRunId}
        asOf={peAsOf}
        onOpenAudit={() => setAuditOpen(true)}
        onLogout={logout}
      />

      {connectionError && (
        <BackendErrorBanner message={connectionError} onDismiss={() => setConnectionError(null)} />
      )}

      <div style={{ display: "flex", flex: 1, minHeight: 0 }}>
        <Sidebar {...sidebarProps} />

        <div style={{ flex: 1, display: "flex", flexDirection: "column", minWidth: 0, background: "var(--bg)" }}>
          <TabBar
            activeTab={activeTab}
            setTab={setActiveTab}
            peCount={peFileState === "loaded" ? peInvestorList.length : 0}
            hfCount={hfFileState === "loaded" ? hfInvestorList.length : 0}
            runStats={peRunId ? { lastRun: `Run ${peRunId.slice(0, 8)}` } : null}
          />

          <div style={{ flex: 1, overflow: "auto" }}>
            {isPe ? (
              <PETab
                fileState={peFileState}
                missingColsError={peMissingCols}
                allInvestors={peInvestorList}
                selectedInvestors={peSelected}
                scope={peScope}
                generationState={peGenerationState}
                generationError={peGenerationError}
                setGenerationState={setPeGenerationState}
                results={peResults}
                setResults={setPeResults}
                insightState={peInsightState}
                insightError={peInsightError}
                setInsightState={setPeInsightState}
                insightMd={peInsightMd}
                setInsightMd={setPeInsightMd}
                insightType={peInsightType}
                setInsightType={setPeInsightType}
                insightScope={peInsightScope}
                setInsightScope={setPeInsightScope}
                insightInvestor={peInsightInvestor}
                setInsightInvestor={setPeInsightInvestor}
                gridView={peGridView}
                setGridView={setPeGridView}
                darkNavyHero={darkNavyHero}
                onTriggerGenerate={handlePeGenerate}
                onTriggerInsight={handlePeInsight}
                downloadUrls={peDownloadUrls}
                runId={peRunId}
                investorsData={pePreviewData}
                partnership={pePartnership}
                asOf={peAsOf}
              />
            ) : (
              <HFTab
                fileState={hfFileState}
                selectedInvestors={hfSelected}
                scope={hfScope}
                generationState={hfGenerationState}
                generationError={hfGenerationError}
                results={hfResults}
                onTriggerGenerate={handleHfGenerate}
                insightState={hfInsightState}
                insightError={hfInsightError}
                insightMd={hfInsightMd}
                insightType={hfInsightType}
                setInsightType={setHfInsightType}
                insightScope={hfInsightScope}
                setInsightScope={setHfInsightScope}
                insightInvestor={hfInsightInvestor}
                setInsightInvestor={setHfInsightInvestor}
                onTriggerInsight={handleHfInsight}
                hasLedger={hfHasLedger}
                onLedgerSelected={handleHfLedgerSelected}
                darkNavyHero={darkNavyHero}
                downloadUrls={hfDownloadUrls}
                runId={hfRunId}
                pcapPreviewData={hfPreviewData}
                fundStats={hfFundStats}
                pcapFileName={hfFileName}
                ledgerFileName={hfLedgerFileName}
              />
            )}
          </div>
        </div>
      </div>

      <ChatDrawer
        open={chatOpen}
        onClose={() => setChatOpen(false)}
        activeTab={activeTab}
        sessionToken={isPe ? peSessionToken : hfPcapToken}
        apiBase={API_BASE}
      />
      <AuditDrawer
        open={auditOpen}
        onClose={() => setAuditOpen(false)}
        apiBase={API_BASE}
      />

      <TweaksPanel title="Tweaks">
        <TweakSection label="Theme">
          <TweakRadio
            label="Hero treatment"
            value={t.theme}
            onChange={v => setTweak("theme", v)}
            options={[
              { value: "light", label: "Light KPMG" },
              { value: "dark",  label: "Dark Navy Hero" },
            ]}
          />
        </TweakSection>
        <TweakSection label="Navigation">
          <TweakRadio
            label="Tab"
            value={activeTab}
            onChange={setActiveTab}
            options={[
              { value: "pe", label: "Private Equity" },
              { value: "hf", label: "Hedge Fund" },
            ]}
          />
        </TweakSection>
      </TweaksPanel>
    </div>
  );
};

const root = ReactDOM.createRoot(document.getElementById("root"));
root.render(<App />);
