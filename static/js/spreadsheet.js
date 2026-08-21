const container = document.getElementById("spreadsheet");

// Worksheet columns passed from Flask
const headers = worksheetColumns;
const cellStyles = {};
// First row = column names
let sheetData;

if (savedWorksheet && savedWorksheet.length > 0) {

    sheetData = savedWorksheet;

} else {

    sheetData = [];
    sheetData.push(headers);

    for (let i = 0; i < 100; i++) {
        sheetData.push(new Array(headers.length).fill(""));
    }

}

// Create Handsontable
const hot = new Handsontable(container, {

    data: sheetData,

    rowHeaders: true,

    colHeaders: false,

    width: "100%",

    height: "75vh",

    stretchH: "all",

    fixedRowsTop: 1,

    fixedColumnsStart: 1,

    contextMenu: true,

    filters: true,

    dropdownMenu: true,

    manualColumnMove: true,

    manualRowMove: true,

    manualColumnResize: true,

    manualRowResize: true,

    licenseKey: "non-commercial-and-evaluation",

    cells: function (row, col) {

        return {

            renderer: function (instance, td, row, col, prop, value, cellProperties) {

                Handsontable.renderers.TextRenderer.apply(this, arguments);

                const key = row + "-" + col;

                if (cellStyles[key]) {

                    Object.assign(td.style, cellStyles[key]);

                }

            }

        };

    },

    

});





// ================= FORMULA BAR =================

let selectedRow = 0;
let selectedCol = 0;

hot.addHook("afterSelection", function(row, col){

    selectedRow = row;
    selectedCol = col;

    const input = document.getElementById("formulaInput");

    if(input){

        input.value = hot.getDataAtCell(row, col);

    }

});

const formulaInput = document.getElementById("formulaInput");

if (formulaInput) {

    formulaInput.addEventListener("keyup", function (e) {

        if (e.key === "Enter") {

            hot.setDataAtCell(selectedRow, selectedCol, this.value);

        }

    });

}



// ================= SAVE =================

function saveWorksheet() {

    // Finish editing the current cell
    hot.getActiveEditor().finishEditing(false);

    // Wait a moment so Handsontable updates the data
    setTimeout(function () {

        const data = hot.getData();

        fetch("/save_worksheet", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                workspace_id: document.getElementById("workspace_id").value,
                data: data
            })
        })
        .then(response => response.json())
        .then(result => {

            console.log(result);
            document.getElementById("saveStatus").innerHTML = "✅ Saved";

        })
        .catch(error => {

            console.error(error);
            document.getElementById("saveStatus").innerHTML = "❌ Save Failed";

        });

    }, 10);
}
let saveTimer = null;

hot.addHook("afterChange", function (changes, source) {

    if (!changes || source === "loadData") return;

    document.getElementById("saveStatus").innerHTML = "🟡 Unsaved Changes";

    clearTimeout(saveTimer);

    saveTimer = setTimeout(function () {

        document.getElementById("saveStatus").innerHTML = "💾 Saving...";

        saveWorksheet();

    }, 2000);

});



// ================= ADD ROW =================

function addRow() {

    const selected = hot.getSelectedLast();

    if (selected) {

        hot.alter("insert_row_below", selected[0]);

    } else {

        hot.alter("insert_row_below", hot.countRows());

    }

}

// ================= ADD COLUMN =================

function addColumn() {

    hot.alter("insert_col_end", 1);

    // Update header row
    hot.setDataAtCell(0, hot.countCols() - 1, "New Column");

}

// ================= DELETE =================

function deleteSelected() {

    const selected = hot.getSelectedLast();

    if (!selected) {

        alert("Select a row first.");

        return;

    }

    // Prevent deleting header row
    if (selected[0] === 0) {

        alert("Header row cannot be deleted.");

        return;

    }

    hot.alter("remove_row", selected[0], 1);

}

// ================= UNDO / REDO =================

function undoWorksheet() {
    hot.undo();
}

function redoWorksheet() {
    hot.redo();
}

// ================= PLACEHOLDERS =================

function analyzeWorksheet() {

    const workspaceId = document.getElementById("workspace_id").value;

    if (!workspaceId) {
        alert("Workspace ID is missing.");
        return;
    }

    window.location.href = "/analyze?workspace_id=" + workspaceId;
}

function openAI() {

    alert("AI Assistant coming soon.");

}


function exportWorksheet() {

    let data = hot.getData();

    // Remove completely empty rows
    data = data.filter(row =>
        row.some(cell => cell !== "" && cell !== null)
    );

    const workbook = XLSX.utils.book_new();

    const worksheet = XLSX.utils.aoa_to_sheet(data);

    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    XLSX.writeFile(workbook, "Worksheet.xlsx");

}
function goWorkspace() {

    window.location.href = "/workspace";

}

let copiedData = "";

function copyCell() {
    copiedData = hot.getDataAtCell(selectedRow, selectedCol);
}

function cutCell() {
    copiedData = hot.getDataAtCell(selectedRow, selectedCol);
    hot.setDataAtCell(selectedRow, selectedCol, "");
}

function pasteCell() {
    if (copiedData !== "") {
        hot.setDataAtCell(selectedRow, selectedCol, copiedData);
    }
}
function getCellKey(){

    return selectedRow + "-" + selectedCol;

}

function boldCell() {

    const key = getCellKey();

    cellStyles[key] = cellStyles[key] || {};

    cellStyles[key].fontWeight =
        cellStyles[key].fontWeight === "bold"
            ? "normal"
            : "bold";

    hot.render();

}

function italicCell() {

    const key = getCellKey();

    cellStyles[key] = cellStyles[key] || {};

    cellStyles[key].fontStyle =
        cellStyles[key].fontStyle === "italic"
            ? "normal"
            : "italic";

    hot.render();

}

function underlineCell() {

    const key = getCellKey();

    cellStyles[key] = cellStyles[key] || {};

    cellStyles[key].textDecoration =
        cellStyles[key].textDecoration === "underline"
            ? "none"
            : "underline";

    hot.render();

}

function changeCellColor(){

    const key = getCellKey();

    cellStyles[key] = cellStyles[key] || {};

    cellStyles[key].background =
    document.getElementById("cellColor").value;

    hot.render();

}

function changeTextColor(){

    const key = getCellKey();

    cellStyles[key] = cellStyles[key] || {};

    cellStyles[key].color =
    document.getElementById("textColor").value;

    hot.render();

}

