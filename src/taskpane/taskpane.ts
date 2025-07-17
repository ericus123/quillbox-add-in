async function handleSelectionChanged(event: Excel.WorksheetSelectionChangedEventArgs) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(event.address);
    range.load("values");
    await context.sync();

    const cellValue = range.values[0][0] || "";
    // Pass cellValue as-is; if it contains HTML, the editor should render it
    openRichTextEditor(cellValue);
  });
}
function openRichTextEditor(cellValue: any) {
  const editorContainer = document.getElementById("rich-text-editor");
  if (editorContainer) {
    editorContainer.innerHTML = cellValue;
    editorContainer.style.display = "block";
  } else {
    console.warn("Rich text editor container not found.");
  }
}
