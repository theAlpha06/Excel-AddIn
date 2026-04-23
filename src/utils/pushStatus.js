export const STEPS = [
  { key: "submitted",  label: "Records submitted" },
  { key: "queued",     label: "Job queued" },
  { key: "processing", label: "Processing quotes" },
  { key: "done",       label: "Salesforce sync complete" },
];

export const BADGE_STYLE = {
  Pending:      { bg: "#FFF4CC", color: "#7A5900" },
  Processing:   { bg: "#E6F1FB", color: "#0C447C" },
  Complete:     { bg: "#EAF3DE", color: "#27500A" },
  Failed:       { bg: "#FCEBEB", color: "#791F1F" },
  RollingBack:  { bg: "#FFF0E0", color: "#8B3A00" },
  RolledBack:   { bg: "#F4E8FF", color: "#4B0082" },
};

export const DOT_COLOR = {
  Pending: "#BA7517", Processing: "#185FA5", Complete: "#3B6D11",
  Failed: "#A32D2D", RollingBack: "#C25A00", RolledBack: "#6A0DAD",
};

export const PROGRESS_COLOR = {
  Pending: "#EF9F27", Processing: "#378ADD", Complete: "#639922",
  Failed: "#E24B4A", RollingBack: "#E2782A", RolledBack: "#9B59B6",
};

export function getProgressPct(status) {
  return {
    Pending: 15, Processing: 55, Complete: 100,
    Failed: 55, RollingBack: 55, RolledBack: 55,
  }[status] ?? 0;
}

export function getStepState(stepKey, status) {
  const order = ["submitted", "queued", "processing", "done"];
  const statusToReached = {
    Pending:      ["submitted", "queued"],
    Processing:   ["submitted", "queued", "processing"],
    Complete:     ["submitted", "queued", "processing", "done"],
    Failed:       ["submitted", "queued", "processing"],
    RollingBack:  ["submitted", "queued", "processing"],
    RolledBack:   ["submitted", "queued", "processing"],
  };

  const reached = statusToReached[status] ?? [];
  const reachedIdx = order.indexOf(reached.at(-1));
  const stepIdx = order.indexOf(stepKey);

  const isErrorStatus = status === "Failed" || status === "RollingBack" || status === "RolledBack";
  if (isErrorStatus && stepKey === "processing") return "error";
  if (stepIdx <= reachedIdx) {
    return status === "Complete" ? "done" : stepIdx === reachedIdx ? "active" : "done";
  }
  return "idle";
}