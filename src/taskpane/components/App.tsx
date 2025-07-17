import React, { useEffect, useState } from "react";
import ReactQuill, { Quill } from "react-quill";
import "react-quill/dist/quill.snow.css";
import "../App.css";
// Office.js API
declare const window: any;

const ExcelQuillEditor: React.FC = () => {
  const [show, setShow] = useState(false);
  const [content, setContent] = useState("");

  // On mount: auto-load selected cell's content and open editor
  useEffect(() => {
    const fetchSelectedCell = async () => {
      try {
        await Excel.run(async (context: any) => {
          const range = context.workbook.getSelectedRange();
          range.load("values");
          await context.sync();
          setContent(range.values[0][0] || "");
        });
      } catch {
        setContent(""); // Fallback if no cell or Office.js fails
      }
      setShow(true);
    };
    fetchSelectedCell();
  }, []);

  const handleSave = async () => {
    await Excel.run(async (context: any) => {
      const range = context.workbook.getSelectedRange();
      range.values = [[content]]; // Save HTML into cell
      await context.sync();
    });
    setShow(false);
  };

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
            boxShadow: "0 4px 24px #0002",
            borderRadius: 10,
          }}
        >
          <ReactQuill
            value={content}
            onChange={setContent}
            style={{
              minHeight: "1000px !important",
            }}
            modules={{
              toolbar: [
                [{ font: [] }],
                [
                  {
                    size: fontSizeArr,
                  },
                ],
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
            formats={[
              "header",
              "font",
              "size",
              "bold",
              "italic",
              "underline",
              "strike",
              "blockquote",
              "code-block",
              "color",
              "background",
              "script",
              "list",
              "bullet",
              "indent",
              "direction",
              "align",
              "link",
            ]}
          />
          <div style={{ marginTop: 16 }}>
            <button
              onClick={handleSave}
              style={{
                background: "#1976d2",
                color: "#fff",
                border: "none",
                borderRadius: 4,
                padding: "8px 20px",
                fontWeight: 500,
                fontSize: 16,
                cursor: "pointer",
                boxShadow: "0 2px 8px #1976d222",
                transition: "background 0.2s",
              }}
            >
              Save
            </button>
            <button
              onClick={() => setShow(false)}
              style={{
                marginLeft: 8,
                background: "#f5f5f5",
                color: "#333",
                border: "1px solid #ccc",
                borderRadius: 4,
                padding: "8px 20px",
                fontWeight: 500,
                fontSize: 16,
                cursor: "pointer",
                boxShadow: "0 2px 8px #0001",
                transition: "background 0.2s, color 0.2s",
              }}
            >
              Cancel
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default ExcelQuillEditor;
