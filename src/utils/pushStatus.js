export const STEPS = [
  { key: "submitted",  label: "Records submitted" },
  { key: "queued",     label: "Job queued" },
  { key: "processing", label: "Processing quotes" },
  { key: "done",       label: "Salesforce sync complete" },
];

export const BADGE_STYLE = {
  Pending:    { bg: "#FFF4CC", color: "#7A5900" },
  Processing: { bg: "#E6F1FB", color: "#0C447C" },
  Complete:   { bg: "#EAF3DE", color: "#27500A" },
  Failed:     { bg: "#FCEBEB", color: "#791F1F" },
};

export const DOT_COLOR = {
  Pending: "#BA7517", Processing: "#185FA5", Complete: "#3B6D11", Failed: "#A32D2D",
};

export const PROGRESS_COLOR = {
  Pending: "#EF9F27", Processing: "#378ADD", Complete: "#639922", Failed: "#E24B4A",
};

export function getProgressPct(status) {
  return { Pending: 15, Processing: 55, Complete: 100, Failed: 55 }[status] ?? 0;
}

export function getStepState(stepKey, status) {
  const order = ["submitted", "queued", "processing", "done"];
  const statusToReached = {
    Pending:    ["submitted", "queued"],
    Processing: ["submitted", "queued", "processing"],
    Complete:   ["submitted", "queued", "processing", "done"],
    Failed:     ["submitted", "queued", "processing"],
  };
  const reached = statusToReached[status] ?? [];
  const reachedIdx = order.indexOf(reached.at(-1));
  const stepIdx = order.indexOf(stepKey);

  if (status === "Failed" && stepKey === "processing") return "error";
  if (stepIdx <= reachedIdx) return status === "Complete" ? "done" : stepIdx === reachedIdx ? "active" : "done";
  return "idle";
}