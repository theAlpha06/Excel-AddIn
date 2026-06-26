export const STEPS = [
  { key: "submitted", label: "Records submitted" },
  { key: "queued", label: "Job queued" },
  { key: "processing", label: "Processing records" },
  { key: "created", label: "Salesforce records created" },
  { key: "writing", label: "Writing IDs to Excel" },
  { key: "uploading", label: "Uploading workbook" },
  { key: "escalations", label: "Updating escalations" },
  { key: "done", label: "Completed" },
];
export const BADGE_STYLE = {
  Pending: {
    bg: "#FFF8E6",
    color: "#9A6700",
  },
  Processing: {
    bg: "#E8F4FD",
    color: "#0B63B6",
  },
  Created: {
    bg: "#E8F4FD",
    color: "#0B63B6",
  },
  Writing: {
    bg: "#E8F4FD",
    color: "#0B63B6",
  },
  Uploading: {
    bg: "#E8F4FD",
    color: "#0B63B6",
  },
  Updating: {
    bg: "#E8F4FD",
    color: "#0B63B6",
  },
  Complete: {
    bg: "#E8F6EC",
    color: "#107C41",
  },
  Failed: {
    bg: "#FDECEC",
    color: "#C42B1C",
  },
  RollingBack: {
    bg: "#FFF4E5",
    color: "#8A4B00",
  },
  RolledBack: {
    bg: "#F3E8FD",
    color: "#6B21A8",
  },
};

export const DOT_COLOR = {
  Pending: "#F4B400",
  Processing: "#0078D4",
  Created: "#0078D4",
  Writing: "#0078D4",
  Uploading: "#0078D4",
  Updating: "#0078D4",
  Complete: "#107C41",
  Failed: "#D13438",
  RollingBack: "#C75C00",
  RolledBack: "#7A3EC8",
};

export const PROGRESS_COLOR = {
  Pending: "#F4B400",
  Processing: "#0078D4",
  Created: "#0078D4",
  Writing: "#0078D4",
  Uploading: "#0078D4",
  Updating: "#0078D4",
  Complete: "#107C41",
  Failed: "#D13438",
  RollingBack: "#C75C00",
  RolledBack: "#7A3EC8",
};

export function getProgressPct(status) {
  return {
    Pending: 10,
    Processing: 30,
    Created: 45,
    Writing: 60,
    Uploading: 75,
    Updating: 90,
    Complete: 100,
    Failed: 30,
    RollingBack: 30,
    RolledBack: 0,
  }[status] ?? 0;
}
export function getStepState(stepKey, status) {
  const order = [
    "submitted",
    "queued",
    "processing",
    "created",
    "writing",
    "uploading",
    "escalations",
    "done",
  ];

  const statusToReached = {
    Pending: ["submitted", "queued"],
    Processing: ["submitted", "queued", "processing"],
    Created: ["submitted", "queued", "processing", "created"],
    Writing: [
      "submitted",
      "queued",
      "processing",
      "created",
      "writing",
    ],
    Uploading: [
      "submitted",
      "queued",
      "processing",
      "created",
      "writing",
      "uploading",
    ],
    Updating: [
      "submitted",
      "queued",
      "processing",
      "created",
      "writing",
      "uploading",
      "escalations",
    ],
    Complete: order,
    Failed: [
      "submitted",
      "queued",
      "processing",
    ],
    RollingBack: [
      "submitted",
      "queued",
      "processing",
    ],
    RolledBack: [
      "submitted",
      "queued",
      "processing",
    ],
  };

  const reached = statusToReached[status] ?? [];
  const reachedIdx = order.indexOf(reached.at(-1));
  const stepIdx = order.indexOf(stepKey);

  const isError =
    status === "Failed" ||
    status === "RollingBack" ||
    status === "RolledBack";

  if (isError && stepIdx === reachedIdx) {
    return "error";
  }

  if (stepIdx < reachedIdx) {
    return "done";
  }

  if (stepIdx === reachedIdx) {
    return status === "Complete" ? "done" : "active";
  }

  return "idle";
}