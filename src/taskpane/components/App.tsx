// ExcelQuillEditor.tsx
import React, { useEffect, useState } from "react";
import ReactQuill, { Quill } from "react-quill";
import "react-quill/dist/quill.snow.css";

declare const Excel: any;
declare const Office: any;

const ExcelQuillEditor: React.FC = () => {
  const [show, setShow] = useState(false);
  const [content, setContent] = useState("");

  // Helper: update editor from current selection
  const loadSelectedCell = async () => {
    try {
      await Excel.run(async (context: any) => {
        const range = context.workbook.getSelectedRange();
        range.load("values");
        await context.sync();
        setContent(range.values[0][0] || "");
      });
    } catch {
      setContent("");
    }
    setShow(true);
  };

  // Run once: listen to Excel selection changes
  useEffect(() => {
    Office.onReady(() => {
      Excel.run(async (context: any) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.onSelectionChanged.add(loadSelectedCell);
        await context.sync();
      });
    });
  }, []);

  const handleSave = async () => {
    await Excel.run(async (context: any) => {
      const range = context.workbook.getSelectedRange();
      range.values = [[content]];
      await context.sync();
    });
    setShow(false);
  };

  // font sizes
  const fontSizeArr = ["8px", "9px", "10px", "12px", "14px", "16px", "20px", "24px", "32px"];
  var Size = Quill.import("attributors/style/size");
  Size.whitelist = fontSizeArr;
  Quill.register(Size, true);

  return (
    <div>
      {show && (
        <div
          style={{
            position: "fixed",
            zIndex: 10,
            margin: "5%",
            background: "#fff",
            border: "1px solid #ddd",
            padding: 24,
            borderRadius: 10,
          }}
        >
          <ReactQuill
            value={content}
            onChange={setContent}
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
          <div style={{ marginTop: 16 }}>
            <button onClick={handleSave}>Save</button>
            <button onClick={() => setShow(false)} style={{ marginLeft: 8 }}>
              Cancel
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default ExcelQuillEditor;
