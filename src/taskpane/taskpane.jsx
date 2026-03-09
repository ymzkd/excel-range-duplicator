/* global Office, Excel */
import React, { useState, useEffect, useCallback } from "react";
import { createRoot } from "react-dom/client";
import {
  FluentProvider,
  webLightTheme,
  Title3,
  Body1,
  Caption1,
  Button,
  SpinButton,
  Field,
  MessageBar,
  MessageBarBody,
  Card,
  CardHeader,
  Divider,
  tokens,
} from "@fluentui/react-components";
import {
  CopyRegular,
  GridRegular,
  TableSimpleRegular,
} from "@fluentui/react-icons";
import "./taskpane.css";

function App() {
  const [selectionAddress, setSelectionAddress] = useState(null);
  const [colCount, setColCount] = useState(2);
  const [rowCount, setRowCount] = useState(1);
  const [status, setStatus] = useState(null); // { message, intent }
  const [running, setRunning] = useState(false);
  const [supported, setSupported] = useState(true);

  // Read current selection from Excel
  const readSelection = useCallback(async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        setSelectionAddress(range.address);
      });
    } catch {
      // Silently ignore transient selection read errors during real-time tracking
    }
  }, []);

  // Setup: check API support + register selection change handler
  useEffect(() => {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
      setSupported(false);
      return;
    }

    // Read initial selection
    readSelection();

    // Register real-time selection change handler
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => readSelection()
    );
  }, [readSelection]);

  // Main duplication logic
  const handleDuplicate = async () => {
    if (!selectionAddress) {
      setStatus({ message: "先に範囲を選択してください。", intent: "error" });
      return;
    }
    if (colCount < 1 || rowCount < 1) {
      setStatus({ message: "正の整数を入力してください。", intent: "error" });
      return;
    }
    if (colCount > 100 || rowCount > 100) {
      setStatus({ message: "各方向の最大複製数は100です。", intent: "error" });
      return;
    }

    setRunning(true);
    setStatus({ message: "複製中...", intent: "info" });

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const sourceRange = sheet.getRange(selectionAddress);
        sourceRange.load(["rowCount", "columnCount"]);
        await context.sync();

        const srcRows = sourceRange.rowCount;
        const srcCols = sourceRange.columnCount;

        for (let r = 0; r < rowCount; r++) {
          for (let c = 0; c < colCount; c++) {
            if (r === 0 && c === 0) continue;

            const targetCell = sourceRange
              .getCell(0, 0)
              .getOffsetRange(r * srcRows, c * srcCols);
            targetCell.copyFrom(sourceRange, Excel.RangeCopyType.all);
          }
        }

        await context.sync();
      });

      const totalCopies = colCount * rowCount - 1;
      setStatus({
        message: `完了! ${colCount} x ${rowCount} のグリッドに${totalCopies}個のコピーを作成しました。`,
        intent: "success",
      });
    } catch (error) {
      setStatus({ message: "エラー: " + error.message, intent: "error" });
    } finally {
      setRunning(false);
    }
  };

  if (!supported) {
    return (
      <div className="app">
        <MessageBar intent="error">
          <MessageBarBody>
            このアドインにはより新しいバージョンの Excel が必要です。
          </MessageBarBody>
        </MessageBar>
      </div>
    );
  }

  return (
    <div className="app">
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <GridRegular fontSize={24} color={tokens.colorBrandForeground1} />
        <Title3>Grid Duplicator</Title3>
      </div>

      <Card size="small">
        <CardHeader
          image={<TableSimpleRegular fontSize={20} />}
          header={<Body1 weight="semibold">選択範囲</Body1>}
          description={
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              リアルタイムで自動取得されます
            </Caption1>
          }
        />
        <div className="selection-box">
          {selectionAddress || "セルを選択してください"}
        </div>
      </Card>

      <Divider />

      <div className="input-group">
        <Field label="横方向の複製数">
          <SpinButton
            value={colCount}
            min={1}
            max={100}
            onChange={(_, data) => {
              if (data.value !== undefined && data.value !== null) {
                setColCount(data.value);
              } else if (data.displayValue !== undefined) {
                const parsed = parseFloat(data.displayValue);
                if (!isNaN(parsed)) {
                  setColCount(Math.min(100, Math.max(1, Math.round(parsed))));
                }
              }
            }}
          />
        </Field>

        <Field label="縦方向の複製数">
          <SpinButton
            value={rowCount}
            min={1}
            max={100}
            onChange={(_, data) => {
              if (data.value !== undefined && data.value !== null) {
                setRowCount(data.value);
              } else if (data.displayValue !== undefined) {
                const parsed = parseFloat(data.displayValue);
                if (!isNaN(parsed)) {
                  setRowCount(Math.min(100, Math.max(1, Math.round(parsed))));
                }
              }
            }}
          />
        </Field>
      </div>

      <Button
        appearance="primary"
        icon={<CopyRegular />}
        disabled={running || !selectionAddress}
        onClick={handleDuplicate}
        size="large"
        style={{ width: "100%" }}
      >
        {running ? "複製中..." : "複製"}
      </Button>

      {status && (
        <MessageBar intent={status.intent}>
          <MessageBarBody>{status.message}</MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
}

// Bootstrap: wait for Office.js, then render React
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const root = createRoot(document.getElementById("root"));
    root.render(
      <FluentProvider theme={webLightTheme}>
        <App />
      </FluentProvider>
    );
  }
});
