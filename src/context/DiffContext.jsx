import React, { createContext, useContext, useState } from "react";

const DiffContext = createContext(null);

export const DiffProvider = ({ children }) => {
  const [pendingPayload, setPendingPayload] = useState(null);
  // pendingPayload shape:
  // {
  //   quoteHeader:    { objectApiName, id, fields } | null,
  //   quoteLineItems: [ { objectApiName, id, fields } ],
  //   quoteCostItems: [ { objectApiName, id, fields } ],
  // }

  const clearPayload = () => setPendingPayload(null);

  return (
    <DiffContext.Provider value={{ pendingPayload, setPendingPayload, clearPayload }}>
      {children}
    </DiffContext.Provider>
  );
};

export const useDiff = () => {
  const ctx = useContext(DiffContext);
  if (!ctx) throw new Error("useDiff must be used within a DiffProvider");
  return ctx;
};