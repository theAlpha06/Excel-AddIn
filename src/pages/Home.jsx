import React, { useState } from 'react';
import Profile from '../components/Profile';
import Progress from '../components/Progress';
import SalesforceService from '../services/SalesforceService';
import { Button } from '@fluentui/react-components';
import { populateSalesforceIds } from '../utils/populateSalesforceIds';

const Home = () => {
  const [activeTab, setActiveTab] = useState("profile");
  const [pushState, setPushState] = useState(null);
  const [isPushing, setIsPushing] = useState(false);

  const handlePush = async () => {
    setIsPushing(true);
    setPushState({ status: "Pending", message: "Submitting…" });

    try {
      await SalesforceService.pushDataToSalesforce(async (update) => {
        const recordErrors = update.results?.filter((r) => r.status === "error") ?? [];
        const hasRecordErrors = recordErrors.length > 0;

        if (update.status === "Complete" && hasRecordErrors) {
          const uniqueMessages = [...new Set(recordErrors.map((r) => r.errorMessage))];
          setPushState({
            ...update,
            status: "Failed",
            message: `Sync completed with ${recordErrors.length} record error(s): ${uniqueMessages.join(", ")}`,
            recordErrors,
          });
          setIsPushing(false);
          return;
        }

        setPushState(update);

        if (update.status === "Complete") {
          setPushState((prev) => ({ ...prev, message: "Writing IDs to Excel…" }));
          await populateSalesforceIds(update);
          setPushState((prev) => ({ ...prev, message: "Done" }));
          setIsPushing(false);
        }

        if (update.status === "Failed") {
          setIsPushing(false);
        }
      });
    } catch (err) {
      setPushState((prev) => ({
        ...prev,
        status: "Failed",
        message: err.message,
      }));
      setIsPushing(false);
    }
  };

  const tabBtn = (key, label) => (
    <div style={{ position: "relative", display: "flex", alignItems: "center" }}>
      <button
        onClick={() => setActiveTab(key)}
        style={{
          padding: "16px 28px",
          border: "none",
          backgroundColor: "transparent",
          color: activeTab === key ? "#0078d4" : "#666",
          cursor: "pointer",
          fontSize: "15px",
          fontWeight: "500",
          position: "relative",
          outline: "none",
        }}
      >
        {label}
        <div
          style={{
            position: "absolute",
            bottom: 0,
            left: 0,
            right: 0,
            height: "3px",
            backgroundColor: "#0078d4",
            transform: activeTab === key ? "scaleX(1)" : "scaleX(0)",
            transformOrigin: activeTab === key ? "left" : "right",
            transition: "transform 0.4s cubic-bezier(0.4, 0, 0.2, 1)",
          }}
        />
      </button>
    </div>
  );

  return (
    <div
      style={{
        width: "100%",
        height: "100%",
        display: "flex",
        flexDirection: "column",
        backgroundColor: "#f9f9f9",
      }}
    >
      <div
        style={{
          display: "flex",
          borderBottom: "1px solid #e5e5e5",
          backgroundColor: "#fff",
          boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
        }}
      >
        {tabBtn("profile", "Profile")}
        {tabBtn("salesforce", "Salesforce Sync")}
      </div>

      <div style={{ flex: 1, overflow: "auto", padding: "10px" }}>
        {activeTab === "profile" && <Profile />}

        {activeTab === "salesforce" && (
          <div style={{ padding: "20px" }}>
            <Button onClick={handlePush} disabled={isPushing}>
              {isPushing
                ? "Pushing…"
                : pushState?.status === "Failed"
                ? "Retry"
                : pushState?.status === "Complete"
                ? "Push Again"
                : "Push"}
            </Button>

            <Progress pushState={pushState} />
          </div>
        )}
      </div>
    </div>
  );
};

export default Home;