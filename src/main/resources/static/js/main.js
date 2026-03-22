const fileInput = document.getElementById("fileInput");
const uploadBtn = document.getElementById("uploadBtn");
const downloadBtn = document.getElementById("downloadBtn");
const resultTable = document.getElementById("resultTable");
const thead = resultTable.querySelector("thead");
const tbody = resultTable.querySelector("tbody");
const messageBox = document.getElementById("message");
const selectedFileName = document.getElementById("selectedFileName");
const rowCount = document.getElementById("rowCount");
const emptyState = document.getElementById("emptyState");
const dropzone = document.getElementById("dropzone");

let currentTableData = [];
let currentToken = null;
let currentOriginalFileName = null;
const changedCellMap = new Map();

fileInput.addEventListener("change", () => {
    const file = fileInput.files[0];
    selectedFileName.textContent = file ? file.name : "선택된 파일 없음";
});

dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
});

dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
});

dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");

    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
        fileInput.files = files;
        selectedFileName.textContent = files[0].name;
    }
});

uploadBtn.addEventListener("click", async () => {
    const file = fileInput.files[0];

    if (!file) {
        setMessage("업로드할 파일을 선택하세요.", "error");
        return;
    }

    const lowerName = file.name.toLowerCase();
    if (!lowerName.endsWith(".xls") && !lowerName.endsWith(".xlsx")) {
        setMessage("엑셀 파일만 업로드할 수 있습니다.", "error");
        return;
    }

    const formData = new FormData();
    formData.append("file", file);

    try {
        setMessage("파일을 업로드하는 중입니다.", "info");

        const response = await fetch("/excel/upload", {
            method: "POST",
            body: formData
        });

        const result = await response.json();

        if (!response.ok) {
            setMessage(result.message || "업로드에 실패했습니다.", "error");
            return;
        }

        currentToken = result.token;
        currentOriginalFileName = result.originalFileName;
        currentTableData = result.rows || [];
        changedCellMap.clear();

        renderTable(currentTableData);

        setMessage("엑셀 데이터를 불러왔습니다. 수식 셀은 수정할 수 없습니다.", "success");
    } catch (e) {
        console.error(e);
        setMessage("서버 호출 중 오류가 발생했습니다.", "error");
    }
});

downloadBtn.addEventListener("click", async () => {
    if (!currentToken) {
        setMessage("먼저 엑셀 파일을 업로드하세요.", "error");
        return;
    }

    const changes = Array.from(changedCellMap.values());

    try {
        setMessage("엑셀 파일을 생성하는 중입니다.", "info");

        const response = await fetch("/excel/download", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                token: currentToken,
                changes: changes
            })
        });

        if (!response.ok) {
            const errorResult = await response.json();
            setMessage(errorResult.message || "다운로드에 실패했습니다.", "error");
            return;
        }

        const blob = await response.blob();
        const downloadUrl = window.URL.createObjectURL(blob);

        const a = document.createElement("a");
        a.href = downloadUrl;
        a.download = buildDownloadFileName(currentOriginalFileName);
        document.body.appendChild(a);
        a.click();
        a.remove();

        window.URL.revokeObjectURL(downloadUrl);

        currentToken = null;
        currentOriginalFileName = null;
        currentTableData = [];
        changedCellMap.clear();
        renderTable([]);

        setMessage("엑셀 파일 다운로드가 완료되었습니다.", "success");
    } catch (e) {
        console.error(e);
        setMessage("다운로드 중 오류가 발생했습니다.", "error");
    }
});

function renderTable(rows) {
    thead.innerHTML = "";
    tbody.innerHTML = "";

    if (!rows || rows.length === 0) {
        rowCount.textContent = "0 rows";
        emptyState.style.display = "flex";
        resultTable.style.display = "none";
        return;
    }

    emptyState.style.display = "none";
    resultTable.style.display = "table";
    rowCount.textContent = `${rows.length} rows`;

    rows.forEach((row, rowVisualIndex) => {
        const tr = document.createElement("tr");

        if (rowVisualIndex === 0) {
            tr.classList.add("excel-first-row");
        }

        row.forEach((cell) => {
            if (cell.hidden) {
                return;
            }

            const td = document.createElement("td");
            td.textContent = cell.value ?? "";

            const rowspan = Number(cell.rowspan || 1);
            const colspan = Number(cell.colspan || 1);

            if (rowspan > 1) {
                td.rowSpan = rowspan;
            }

            if (colspan > 1) {
                td.colSpan = colspan;
            }

            td.dataset.rowIndex = cell.rowIndex;
            td.dataset.colIndex = cell.colIndex;
            td.dataset.rowspan = rowspan;
            td.dataset.colspan = colspan;
            td.dataset.formulaCell = String(!!cell.formulaCell);

            if (cell.formulaCell) {
                td.contentEditable = "false";
                td.classList.add("formula-cell");
                td.title = "수식 셀은 수정할 수 없습니다.";
            } else {
                td.contentEditable = "true";
                td.classList.add("editable-cell");

                td.addEventListener("input", (e) => {
                    const newValue = e.target.textContent ?? "";
                    cell.value = newValue;

                    const key = `${cell.rowIndex}:${cell.colIndex}`;
                    changedCellMap.set(key, {
                        rowIndex: Number(cell.rowIndex),
                        colIndex: Number(cell.colIndex),
                        value: newValue
                    });
                });
            }

            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });
}

function buildDownloadFileName(originalFileName) {
    if (!originalFileName) {
        return "edited_excel.xlsx";
    }
    return `edited_${originalFileName}`;
}

function setMessage(text, type = "info") {
    messageBox.textContent = text;
    messageBox.className = `message-box ${type}`;
}