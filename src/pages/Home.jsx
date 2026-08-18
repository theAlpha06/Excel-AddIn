import React, { useEffect, useState } from 'react';
import Profile from '../components/Profile';
import Progress from '../components/Progress';
import SalesforceService from '../services/SalesforceService';
import { Button } from '@fluentui/react-components';
import { populateSalesforceIds } from '../utils/populateSalesforceIds';
import SalesforceSync from './SalesforceSync';
import { divProperties } from '@fluentui/react';

const Home = () => {
  const [activeTab, setActiveTab] = useState("profile");

  const lockFields = async () => {
    await Excel.run({ delayForCellEdit: true }, async (context) => {
      const sheet1 = context.workbook.worksheets.getItem(
        "Final Quote Header & Lines"
      );

      const worksheets = context.workbook.worksheets;
      worksheets.load("items/name");

      sheet1.protection.load("protected");

      await context.sync();

      const lineTemplateSheets = worksheets.items.filter(sheet =>
        sheet.name.startsWith("Template")
      );

      lineTemplateSheets.forEach(sheet => {
        sheet.protection.load("protected");
      });

      await context.sync();

      if (sheet1.protection.protected) {
        sheet1.protection.unprotect();
      }

      lineTemplateSheets.forEach(sheet => {
        if (sheet.protection.protected) {
          sheet.protection.unprotect();
        }
      });

      sheet1.getUsedRange().format.protection.locked = false;
      sheet1.getRange("A4:A18").format.protection.locked = true;
      sheet1.getRange("A20").format.protection.locked = true;
      sheet1.getRange("A21:I21").format.protection.locked = true;
      sheet1.getRange("G6:G13").format.protection.locked = true;
      sheet1.getRange("G15:L15").format.protection.locked = true;
      sheet1.getRange("G5:L5").format.protection.locked = true;
      sheet1.getRange("G14:I14").format.protection.locked = true;

      lineTemplateSheets.forEach(sheet => {
        sheet.getUsedRange().format.protection.locked = false;
        sheet.getRange("A5:D5").format.protection.locked = true;
      });

      sheet1.protection.protect({
        allowAutoFilter: true,
        allowSort: true,
        allowInsertRows: true,
      });

      lineTemplateSheets.forEach(sheet => {
        sheet.protection.protect({
          allowAutoFilter: true,
          allowSort: true,
          allowInsertRows: true,
        });
      });

      await context.sync();
    });
  };

  useEffect(() => {
    lockFields();
  }, [])

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

        {activeTab === "salesforce" && <SalesforceSync />}
      </div>
    </div>
  );
};

export default Home;