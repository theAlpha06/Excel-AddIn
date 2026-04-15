import React from 'react';
import {
  STEPS, BADGE_STYLE, DOT_COLOR, PROGRESS_COLOR,
  getProgressPct, getStepState,
} from '../utils/pushStatus';

const STEP_COLOR = {
  done:   { bg: "#EAF3DE", color: "#3B6D11" },
  active: { bg: "#E6F1FB", color: "#185FA5" },
  error:  { bg: "#FCEBEB", color: "#A32D2D" },
  idle:   { bg: "#f5f5f5", color: "#aaa" },
};

const STEP_ICON = { done: "✓", active: "…", error: "✕" };

const Progress = ({ pushState }) => {
  if (!pushState) return null;

  const { status, jobKey, message, recordErrors } = pushState;
  const badgeStyle = BADGE_STYLE[status] ?? {};
  const progressPct = getProgressPct(status);

  return (
    <div style={{
      marginTop: "16px", border: "0.5px solid #e5e5e5",
      borderRadius: "12px", padding: "16px 20px", background: "#fff",
    }}>

      <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "14px" }}>
        <span style={{ fontSize: "13px", color: "#888" }}>Status</span>
        <span style={{
          display: "inline-flex", alignItems: "center", gap: "6px",
          fontSize: "12px", fontWeight: 500, padding: "3px 10px",
          borderRadius: "99px", background: badgeStyle.bg, color: badgeStyle.color,
        }}>
          <span style={{
            width: "7px", height: "7px", borderRadius: "50%",
            background: DOT_COLOR[status],
            animation: status === "Processing" ? "sf-pulse 1.2s ease-in-out infinite" : "none",
          }} />
          {status}
        </span>

        {message && (
          <span style={{ fontSize: "12px", color: "#888", marginLeft: "4px" }}>
            {message}
          </span>
        )}
      </div>

      <div style={{ height: "4px", background: "#f0f0f0", borderRadius: "99px", overflow: "hidden", marginBottom: "16px" }}>
        <div style={{
          height: "100%", borderRadius: "99px",
          width: `${progressPct}%`,
          background: PROGRESS_COLOR[status],
          transition: "width 0.6s ease",
        }} />
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
        {STEPS.map(({ key, label }, i) => {
          const state = getStepState(key, status);
          const c = STEP_COLOR[state];
          const icon = STEP_ICON[state] ?? String(i + 1);
          const isError = state === "error" && key === "processing";
          return (
            <div key={key} style={{ display: "flex", alignItems: "center", gap: "10px", fontSize: "13px" }}>
              <div style={{
                width: "20px", height: "20px", borderRadius: "50%", flexShrink: 0,
                display: "flex", alignItems: "center", justifyContent: "center",
                fontSize: "10px", background: c.bg, color: c.color,
              }}>
                {icon}
              </div>
              <span style={{
                color: state === "idle" ? "#aaa" : state === "error" ? "#A32D2D" : "#111",
                fontWeight: state === "active" ? 500 : 400,
              }}>
                {isError ? "Error: processing failed" : label}
              </span>
            </div>
          );
        })}
      </div>

      {recordErrors?.length > 0 && (
        <div style={{
          marginTop: "14px",
          paddingTop: "12px",
          borderTop: "0.5px solid #eee",
        }}>
          <div style={{
            fontSize: "12px", fontWeight: 500, color: "#A32D2D", marginBottom: "8px",
          }}>
            {recordErrors.length} record{recordErrors.length > 1 ? "s" : ""} failed to sync
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: "5px" }}>
            {recordErrors.map((err) => (
              <div
                key={err.rowIndex}
                style={{
                  display: "flex", alignItems: "center", justifyContent: "space-between",
                  fontSize: "12px", padding: "6px 10px",
                  background: "#FCEBEB", borderRadius: "8px",
                }}
              >
                <span style={{ color: "#791F1F", fontFamily: "monospace" }}>
                  Row {err.rowIndex} &nbsp;·&nbsp; {err.cell}
                </span>
                <span style={{ color: "#A32D2D", fontWeight: 500 }}>
                  {err.errorMessage}
                </span>
              </div>
            ))}
          </div>
        </div>
      )}

      {jobKey && (
        <div style={{
          marginTop: "14px", paddingTop: "12px",
          borderTop: "0.5px solid #eee",
          fontSize: "11px", color: "#aaa", fontFamily: "monospace",
        }}>
          Job: {jobKey}
          {status === "Complete"  && " · synced successfully"}
          {status === "Failed"    && " · check Salesforce logs"}
          {(status === "Pending" || status === "Processing") && " · polling every 3s"}
        </div>
      )}

      <style>{`@keyframes sf-pulse { 0%,100%{opacity:1} 50%{opacity:0.3} }`}</style>
    </div>
  );
};

export default Progress;