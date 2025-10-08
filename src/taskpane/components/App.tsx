import React, { useEffect, useState } from "react";
import ReactQuill, { Quill } from "react-quill";
import "react-quill/dist/quill.snow.css";
import "../App.css";

declare const Office: any;

const ExcelQuillEditor: React.FC = () => {
  const [content, setContent] = useState("");
  const [excelReady, setExcelReady] = useState(false);
  const [viewMode, setViewMode] = useState<"editor" | "code" | "preview">("editor");

  // --- Setup Quill font sizes
  const fontSizeArr = [
    "8px",
    "9px",
    "10px",
    "12px",
    "14px",
    "16px",
    "20px",
    "24px",
    "32px",
    "40px",
    "48px",
    "56px",
    "64px",
    "72px",
    "96px",
    "128px",
  ];
  const Size = Quill.import("attributors/style/size");
  Size.whitelist = fontSizeArr;
  Quill.register(Size, true);

  // --- Load selected Excel cell
  const loadSelectedCell = async () => {
    const Excel = (window as any).Excel;
    if (!Excel) return; // Not in Excel
    try {
      await Excel.run(async (context: any) => {
        const range = context.workbook.getSelectedRange();
        range.load("values");
        await context.sync();
        setContent(range.values?.[0]?.[0] || "");
      });
    } catch (err) {
      console.warn("Excel read failed:", err);
    }
  };

  // --- Save back to Excel cell
  const handleSave = async () => {
    const Excel = (window as any).Excel;
    if (!Excel) {
      alert("Excel environment not detected — cannot save to cell.");
      return;
    }
    try {
      await Excel.run(async (context: any) => {
        const range = context.workbook.getSelectedRange();
        range.values = [[content]];
        await context.sync();
      });
    } catch (err) {
      console.warn("Excel write failed:", err);
    }
  };

  // --- Detect Excel environment
  useEffect(() => {
    if (typeof Office !== "undefined" && Office.onReady) {
      Office.onReady((info: any) => {
        if (info.host === Office.HostType.Excel) {
          setExcelReady(true);
          setTimeout(() => {
            const Excel = (window as any).Excel;
            if (!Excel) return;
            Excel.run(async (context: any) => {
              const sheet = context.workbook.worksheets.getActiveWorksheet();
              sheet.onSelectionChanged.add(loadSelectedCell);
              await context.sync();
            });
          }, 400);
        }
      });
    } else {
      console.log("Office.js not detected — standalone mode.");
    }
  }, []);

  // --- UI
  return (
    <div
      style={{
        background: "#f9f9fb",
        borderRadius: "12px",
        padding: "20px",
        margin: "40px auto",
        boxShadow: "0 4px 14px rgba(0,0,0,0.1)",
        maxWidth: "900px",
        width: "90%",
        fontFamily: "Inter, sans-serif",
      }}
    >
      <header
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          marginBottom: "15px",
        }}
      >
        <h2 style={{ margin: 0, fontWeight: 600, color: "#333" }}>
          Excel WYSIWYG Editor{" "}
          <span style={{ fontSize: "14px", color: excelReady ? "green" : "#888" }}>
            {excelReady ? "• Connected" : "• Standalone"}
          </span>
        </h2>
        <div style={{ display: "flex", gap: "8px" }}>
          <button
            onClick={loadSelectedCell}
            style={buttonStyle("green")}
            title="Load selected cell"
          >
            ⟳ Load
          </button>
          <button onClick={handleSave} style={buttonStyle("#0078d4")} title="Save to Excel">
            💾 Save
          </button>
          <button
            onClick={() => setViewMode("editor")}
            style={buttonStyle(viewMode === "editor" ? "#555" : "#ccc")}
            title="Edit mode"
          >
            ✏️ Edit
          </button>
          <button
            onClick={() => setViewMode("code")}
            style={buttonStyle(viewMode === "code" ? "#555" : "#ccc")}
            title="View HTML"
          >
            {"</>"}
          </button>
          <button
            onClick={() => setViewMode("preview")}
            style={buttonStyle(viewMode === "preview" ? "#555" : "#ccc")}
            title="Preview content"
          >
            👁 Preview
          </button>
        </div>
      </header>

      {/* --- Content Modes --- */}
      <div
        style={{
          border: "1px solid #ddd",
          borderRadius: "8px",
          background: "#fff",
          overflow: "hidden",
          minHeight: "300px",
        }}
      >
        {viewMode === "editor" && (
          <ReactQuill
            theme="snow"
            value={content}
            onChange={setContent}
            style={{ height: "400px" }}
            modules={{
              toolbar: [
                [{ font: [] }],
                [{ size: fontSizeArr }],
                [{ header: [1, 2, 3, 4, 5, 6, false] }],
                ["bold", "italic", "underline", "strike"],
                [{ color: [] }, { background: [] }],
                [{ list: "ordered" }, { list: "bullet" }, { indent: "-1" }, { indent: "+1" }],
                [{ direction: "rtl" }],
                [{ align: [] }],
                ["blockquote", "code-block"],
                ["link"],
                ["clean"],
              ],
            }}
          />
        )}

        {viewMode === "code" && (
          <div
            style={{
              padding: "16px",
              fontFamily: "monospace",
              whiteSpace: "pre-wrap",
              background: "#111",
              color: "#00ff88",
              fontSize: "14px",
              height: "400px",
              overflowY: "auto",
            }}
          >
            {content || "<p><i>No content yet...</i></p>"}
          </div>
        )}

        {viewMode === "preview" && (
          <div
            style={{
              padding: "20px",
              height: "400px",
              overflowY: "auto",
              background: "#fff",
            }}
            dangerouslySetInnerHTML={{ __html: content || "<p><i>No content yet...</i></p>" }}
          />
        )}
      </div>
    </div>
  );
};

// Helper for button styling
function buttonStyle(color: string): React.CSSProperties {
  return {
    backgroundColor: color,
    border: "none",
    borderRadius: "6px",
    color: color === "#ccc" ? "#333" : "#fff",
    fontWeight: 600,
    padding: "8px 14px",
    cursor: "pointer",
    transition: "all 0.2s ease",
  };
}

export default ExcelQuillEditor;
