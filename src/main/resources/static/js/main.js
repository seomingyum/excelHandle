let currentExcelData = null;

document.addEventListener("DOMContentLoaded", function () {
    const uploadBtn = document.getElementById("uploadBtn");
    const downloadBtn = document.getElementById("downloadBtn");
    const fileInput = document.getElementById("fileInput");
    const dropzone = document.getElementById("dropzone");

    uploadBtn.addEventListener("click", uploadExcel);
    downloadBtn.addEventListener("click", downloadExcelFromTable);
    fileInput.addEventListener("change", handleFileSelect);

    initDropzone(dropzone, fileInput);
    updateTableState(false);
    updateRowCount(0);
});

function handleFileSelect() {
    const fileInput = document.getElementById("fileInput");
    const selectedFileName = document.getElementById("selectedFileName");
    const file = fileInput.files[0];

    selectedFileName.textContent = file ? file.name : "선택된 파일 없음";

    if (file) {
        setMessage("파일이 선택되었습니다. 외부에서 수정했다면 다시 선택 후 업로드하세요.", "info");
    }
}

function initDropzone(dropzone, fileInput) {
    if (!dropzone) {
        return;
    }

    ["dragenter", "dragover"].forEach(function (eventName) {
        dropzone.addEventListener(eventName, function (e) {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.add("dragover");
        });
    });

    ["dragleave", "drop"].forEach(function (eventName) {
        dropzone.addEventListener(eventName, function (e) {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove("dragover");
        });
    });

    dropzone.addEventListener("drop", function (e) {
        const files = e.dataTransfer.files;
        if (files && files.length > 0) {
            fileInput.files = files;
            handleFileSelect();
        }
    });
}

async function uploadExcel() {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput.files[0];

    clearTable();
    updateTableState(false);
    updateRowCount(0);
    currentExcelData = null;

    if (!file) {
        setMessage("파일을 선택하세요.", "error");
        return;
    }

    setMessage("파일 업로드 중입니다.", "info");

    const formData = new FormData();
    formData.append("file", file);

    try {
        const response = await fetch("/excel/upload", {
            method: "POST",
            body: formData
        });

        const data = await response.json();

        if (!response.ok) {
            setMessage(data.message || "업로드 중 오류가 발생했습니다.", "error");
            return;
        }

        renderEditableTable(data);
        setMessage("업로드 및 파싱이 완료되었습니다.", "success");
    } catch (e) {
        console.error("uploadExcel fetch error:", e);
        setMessage(
            "서버 호출 중 오류가 발생했습니다. 파일을 저장한 뒤 엑셀을 닫고 다시 선택 후 업로드하세요.",
            "error"
        );
    }
}

function renderEditableTable(data) {
    currentExcelData = JSON.parse(JSON.stringify(data));

    const thead = document.querySelector("#resultTable thead");
    const tbody = document.querySelector("#resultTable tbody");

    thead.innerHTML = "";
    tbody.innerHTML = "";

    if (!data.headers || data.headers.length === 0) {
        updateTableState(false);
        updateRowCount(0);
        return;
    }

    data.headers.forEach(function (headerRow, rowIndex) {
        const tr = document.createElement("tr");

        headerRow.forEach(function (cell, colIndex) {
            if (cell.hidden) {
                return;
            }

            const th = document.createElement("th");
            th.textContent = cell.value ?? "";
            th.contentEditable = "true";
            th.dataset.section = "header";
            th.dataset.rowIndex = rowIndex;
            th.dataset.colIndex = colIndex;

            if (cell.rowspan && cell.rowspan > 1) {
                th.rowSpan = cell.rowspan;
            }

            if (cell.colspan && cell.colspan > 1) {
                th.colSpan = cell.colspan;
            }

            tr.appendChild(th);
        });

        thead.appendChild(tr);
    });

    if (!data.rows || data.rows.length === 0) {
        const tr = document.createElement("tr");
        const td = document.createElement("td");
        td.colSpan = getVisibleColumnCount(data.headers[0]);
        td.textContent = "데이터가 없습니다.";
        tr.appendChild(td);
        tbody.appendChild(tr);

        updateTableState(true);
        updateRowCount(0);
        return;
    }

    data.rows.forEach(function (row, rowIndex) {
        const tr = document.createElement("tr");

        row.forEach(function (cell, colIndex) {
            if (cell.hidden) {
                return;
            }

            const td = document.createElement("td");
            td.textContent = cell.value ?? "";
            td.contentEditable = "true";
            td.dataset.section = "body";
            td.dataset.rowIndex = rowIndex;
            td.dataset.colIndex = colIndex;

            if (cell.rowspan && cell.rowspan > 1) {
                td.rowSpan = cell.rowspan;
            }

            if (cell.colspan && cell.colspan > 1) {
                td.colSpan = cell.colspan;
            }

            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });

    updateTableState(true);
    updateRowCount(data.rows.length);
}

function downloadExcelFromTable() {
    if (!currentExcelData) {
        setMessage("다운로드할 테이블 데이터가 없습니다.", "error");
        return;
    }

    const clonedData = JSON.parse(JSON.stringify(currentExcelData));

    applyEditedValuesToData(clonedData);

    const sheetData = buildSheetDataFromExcelData(clonedData);
    const merges = buildMergesFromExcelData(clonedData);

    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    worksheet["!merges"] = merges;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    XLSX.writeFile(workbook, "edited_excel.xlsx");

    setMessage("병합 정보가 유지된 엑셀 파일 다운로드가 완료되었습니다.", "success");
}

function applyEditedValuesToData(excelData) {
    const editableCells = document.querySelectorAll(
        "#resultTable th[contenteditable='true'], #resultTable td[contenteditable='true']"
    );

    editableCells.forEach(function (el) {
        const section = el.dataset.section;
        const rowIndex = Number(el.dataset.rowIndex);
        const colIndex = Number(el.dataset.colIndex);
        const value = el.textContent ?? "";

        if (section === "header") {
            excelData.headers[rowIndex][colIndex].value = value;
        } else if (section === "body") {
            excelData.rows[rowIndex][colIndex].value = value;
        }
    });
}

function buildSheetDataFromExcelData(excelData) {
    const result = [];

    excelData.headers.forEach(function (row) {
        const rowData = [];
        row.forEach(function (cell) {
            rowData.push(cell.hidden ? "" : (cell.value ?? ""));
        });
        result.push(rowData);
    });

    excelData.rows.forEach(function (row) {
        const rowData = [];
        row.forEach(function (cell) {
            rowData.push(cell.hidden ? "" : (cell.value ?? ""));
        });
        result.push(rowData);
    });

    return result;
}

function buildMergesFromExcelData(excelData) {
    const merges = [];
    const headerRowCount = excelData.headers.length;

    excelData.headers.forEach(function (row, rowIndex) {
        row.forEach(function (cell, colIndex) {
            if (cell.hidden) {
                return;
            }

            const rowspan = cell.rowspan || 1;
            const colspan = cell.colspan || 1;

            if (rowspan > 1 || colspan > 1) {
                merges.push({
                    s: { r: rowIndex, c: colIndex },
                    e: { r: rowIndex + rowspan - 1, c: colIndex + colspan - 1 }
                });
            }
        });
    });

    excelData.rows.forEach(function (row, rowIndex) {
        row.forEach(function (cell, colIndex) {
            if (cell.hidden) {
                return;
            }

            const rowspan = cell.rowspan || 1;
            const colspan = cell.colspan || 1;

            if (rowspan > 1 || colspan > 1) {
                merges.push({
                    s: { r: headerRowCount + rowIndex, c: colIndex },
                    e: {
                        r: headerRowCount + rowIndex + rowspan - 1,
                        c: colIndex + colspan - 1
                    }
                });
            }
        });
    });

    return merges;
}

function getVisibleColumnCount(row) {
    let count = 0;

    row.forEach(function (cell) {
        if (!cell.hidden) {
            count += cell.colspan && cell.colspan > 1 ? cell.colspan : 1;
        }
    });

    return count;
}

function clearTable() {
    document.querySelector("#resultTable thead").innerHTML = "";
    document.querySelector("#resultTable tbody").innerHTML = "";
}

function updateTableState(hasData) {
    const table = document.getElementById("resultTable");
    const emptyState = document.getElementById("emptyState");

    if (hasData) {
        table.style.display = "table";
        emptyState.style.display = "none";
    } else {
        table.style.display = "none";
        emptyState.style.display = "flex";
    }
}

function updateRowCount(count) {
    document.getElementById("rowCount").textContent = count + " rows";
}

function setMessage(text, type) {
    const message = document.getElementById("message");
    message.textContent = text;
    message.className = "message-box " + type;
}