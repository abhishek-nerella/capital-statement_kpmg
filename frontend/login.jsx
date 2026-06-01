/* KPMG Capital Analysis Statement Generator — Login */

// ── Credentials — edit these two values ───────────────────────────────────────
const CREDENTIALS = {
  username: "kpmg-poc",
  password: "PeHF@1234",
};

const AUTH_SESSION_KEY = "kpmg_casg_auth";
const AUTH_LOCAL_KEY   = "kpmg_casg_auth_persist";

// ── Auth hook — shared across the whole app ───────────────────────────────────
const useAuth = () => {
  const [isAuthenticated, setIsAuthenticated] = React.useState(() =>
    sessionStorage.getItem(AUTH_SESSION_KEY) === "1" ||
    localStorage.getItem(AUTH_LOCAL_KEY)   === "1"
  );

  const login = (username, password, remember) => {
    if (username.trim() === CREDENTIALS.username &&
        password       === CREDENTIALS.password) {
      sessionStorage.setItem(AUTH_SESSION_KEY, "1");
      if (remember) localStorage.setItem(AUTH_LOCAL_KEY, "1");
      setIsAuthenticated(true);
      return true;
    }
    return false;
  };

  const logout = () => {
    sessionStorage.removeItem(AUTH_SESSION_KEY);
    localStorage.removeItem(AUTH_LOCAL_KEY);
    setIsAuthenticated(false);
  };

  return { isAuthenticated, login, logout };
};

// ── Login page ────────────────────────────────────────────────────────────────
const LoginPage = ({ onLogin }) => {
  const [username, setUsername] = React.useState("");
  const [password, setPassword] = React.useState("");
  const [remember, setRemember] = React.useState(false);
  const [showPw,   setShowPw]   = React.useState(false);
  const [error,    setError]    = React.useState("");
  const [loading,  setLoading]  = React.useState(false);
  const [shake,    setShake]    = React.useState(false);

  const canSubmit = username.trim().length > 0 && password.length > 0 && !loading;

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!canSubmit) return;
    setLoading(true);
    setError("");

    // Small UX delay so the spinner is perceptible
    setTimeout(() => {
      const ok = onLogin(username.trim(), password, remember);
      if (!ok) {
        setError("Incorrect username or password.");
        setLoading(false);
        setShake(true);
        setTimeout(() => setShake(false), 600);
      }
    }, 380);
  };

  const inputStyle = (focused) => ({
    width: "100%",
    padding: "10px 13px",
    border: `1.5px solid ${focused ? "#00338D" : "#d1d9e0"}`,
    borderRadius: 3,
    fontSize: 13.5,
    fontFamily: "var(--font-sans, 'Inter', sans-serif)",
    color: "#0a1830",
    outline: "none",
    background: "#fff",
    boxSizing: "border-box",
    transition: "border-color 0.15s",
  });

  // Individual input focus state
  const [uFocus, setUFocus] = React.useState(false);
  const [pFocus, setPFocus] = React.useState(false);

  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(150deg, #0C233C 0%, #00338D 55%, #1E49E2 100%)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontFamily: "var(--font-sans, 'Inter', sans-serif)",
      padding: 20,
    }}>

      {/* Login card */}
      <div style={{
        background: "#fff",
        width: "100%",
        maxWidth: 420,
        boxShadow: "0 32px 80px rgba(0,0,0,0.40)",
        transform: shake ? "translateX(0)" : "none",
        animation: shake ? "shake 0.55s ease" : "none",
      }}>

        {/* Pacific accent strip */}
        <div style={{ height: 4, background: "#00B8F5" }} />

        {/* Blue header */}
        <div style={{
          background: "#00338D",
          padding: "30px 40px 26px",
          textAlign: "center",
        }}>
          <div style={{
            fontSize: 34,
            fontWeight: 900,
            color: "#fff",
            letterSpacing: "0.07em",
            lineHeight: 1,
            marginBottom: 8,
          }}>
            KPMG
          </div>
          <div style={{
            fontSize: 11.5,
            color: "#ACEAFF",
            fontWeight: 500,
            letterSpacing: "0.10em",
            textTransform: "uppercase",
          }}>
            Capital Analysis Statement Generator
          </div>
        </div>

        {/* Form area */}
        <div style={{ padding: "32px 40px 28px" }}>
          <div style={{
            fontSize: 18,
            fontWeight: 700,
            color: "#00338D",
            marginBottom: 4,
          }}>
            Sign in
          </div>
          <div style={{
            fontSize: 12.5,
            color: "#64748b",
            marginBottom: 26,
          }}>
            Authorised users only · For internal use
          </div>

          <form onSubmit={handleSubmit} autoComplete="on">

            {/* Username */}
            <div style={{ marginBottom: 16 }}>
              <label style={{
                display: "block",
                fontSize: 11,
                fontWeight: 700,
                color: "#00338D",
                textTransform: "uppercase",
                letterSpacing: "0.09em",
                marginBottom: 5,
              }}>
                Username
              </label>
              <input
                type="text"
                value={username}
                onChange={e => setUsername(e.target.value)}
                autoFocus
                autoComplete="username"
                placeholder="Enter your username"
                style={inputStyle(uFocus)}
                onFocus={() => setUFocus(true)}
                onBlur={()  => setUFocus(false)}
              />
            </div>

            {/* Password */}
            <div style={{ marginBottom: 20 }}>
              <label style={{
                display: "block",
                fontSize: 11,
                fontWeight: 700,
                color: "#00338D",
                textTransform: "uppercase",
                letterSpacing: "0.09em",
                marginBottom: 5,
              }}>
                Password
              </label>
              <div style={{ position: "relative" }}>
                <input
                  type={showPw ? "text" : "password"}
                  value={password}
                  onChange={e => setPassword(e.target.value)}
                  autoComplete="current-password"
                  placeholder="Enter your password"
                  style={{ ...inputStyle(pFocus), paddingRight: 42 }}
                  onFocus={() => setPFocus(true)}
                  onBlur={()  => setPFocus(false)}
                />
                <button
                  type="button"
                  tabIndex={-1}
                  onClick={() => setShowPw(v => !v)}
                  style={{
                    position: "absolute",
                    right: 11,
                    top: "50%",
                    transform: "translateY(-50%)",
                    background: "none",
                    border: "none",
                    cursor: "pointer",
                    color: "#94a3b8",
                    padding: 3,
                    display: "flex",
                    alignItems: "center",
                  }}
                  title={showPw ? "Hide password" : "Show password"}
                >
                  {showPw ? (
                    /* eye-off */
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none"
                         stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94"/>
                      <path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19"/>
                      <line x1="1" y1="1" x2="23" y2="23"/>
                    </svg>
                  ) : (
                    /* eye */
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none"
                         stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                      <circle cx="12" cy="12" r="3"/>
                    </svg>
                  )}
                </button>
              </div>
            </div>

            {/* Remember me */}
            <div style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              marginBottom: 22,
              cursor: "pointer",
            }}
              onClick={() => setRemember(v => !v)}
            >
              <div style={{
                width: 15,
                height: 15,
                border: `2px solid ${remember ? "#00338D" : "#cbd5e1"}`,
                background: remember ? "#00338D" : "#fff",
                borderRadius: 2,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                flexShrink: 0,
                transition: "all 0.15s",
              }}>
                {remember && (
                  <svg width="9" height="9" viewBox="0 0 12 12" fill="none"
                       stroke="#fff" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                    <polyline points="2 6 5 9 10 3"/>
                  </svg>
                )}
              </div>
              <span style={{ fontSize: 12.5, color: "#475569", userSelect: "none" }}>
                Keep me signed in
              </span>
            </div>

            {/* Error banner */}
            {error && (
              <div style={{
                display: "flex",
                alignItems: "center",
                gap: 9,
                background: "#fff1f2",
                border: "1px solid #fecdd3",
                borderLeft: "3px solid #e11d48",
                borderRadius: 3,
                padding: "9px 13px",
                marginBottom: 18,
                fontSize: 12.5,
                color: "#be123c",
              }}>
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none"
                     stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="12" y1="8" x2="12" y2="12"/>
                  <line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                {error}
              </div>
            )}

            {/* Submit button */}
            <button
              type="submit"
              disabled={!canSubmit}
              style={{
                width: "100%",
                padding: "12px 0",
                background: canSubmit ? "#00338D" : "#cbd5e1",
                color: canSubmit ? "#fff" : "#94a3b8",
                border: "none",
                borderRadius: 3,
                fontSize: 13.5,
                fontWeight: 700,
                letterSpacing: "0.02em",
                cursor: canSubmit ? "pointer" : "not-allowed",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                gap: 8,
                transition: "background 0.15s",
              }}
              onMouseEnter={e => { if (canSubmit) e.currentTarget.style.background = "#002270"; }}
              onMouseLeave={e => { if (canSubmit) e.currentTarget.style.background = "#00338D"; }}
            >
              {loading ? (
                <>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none"
                       stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"
                       style={{ animation: "spin 0.8s linear infinite" }}>
                    <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
                  </svg>
                  Signing in…
                </>
              ) : (
                <>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none"
                       stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/>
                    <polyline points="10 17 15 12 10 7"/>
                    <line x1="15" y1="12" x2="3" y2="12"/>
                  </svg>
                  Sign In
                </>
              )}
            </button>
          </form>
        </div>

        {/* Footer */}
        <div style={{
          borderTop: "1px solid #f1f5f9",
          padding: "13px 40px 18px",
          textAlign: "center",
          fontSize: 11,
          color: "#94a3b8",
          lineHeight: 1.6,
        }}>
          © KPMG International · Confidential<br />
          Authorised users only · Unauthorised access is prohibited
        </div>
      </div>

      {/* CSS animations injected inline */}
      <style>{`
        @keyframes spin  { to { transform: rotate(360deg); } }
        @keyframes shake {
          0%,100% { transform: translateX(0); }
          18%     { transform: translateX(-7px); }
          36%     { transform: translateX(7px); }
          54%     { transform: translateX(-5px); }
          72%     { transform: translateX(5px); }
          90%     { transform: translateX(-2px); }
        }
      `}</style>
    </div>
  );
};

window.useAuth    = useAuth;
window.LoginPage  = LoginPage;
