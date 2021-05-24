//------------------ADDING PERFECT SCROLLBAR TO CELLS------------------//

const ps = new PerfectScrollbar("#cells", {
    wheelPropagation: true
});

//------------------GLOBAL VARIABLES------------------//

let saved = true;
let selectedSheet = "Sheet1";
let lastAddedSheet = 1;
let totalSheets = 1;
let startCellSelected = false;
let startCell = {};
let endCell = {};
let clipboard = { startCell: [], cellData: {} };
let cutClicked = false;

let scrollRightStarted = false;
let scrollLeftStarted = false;
let scrollTopStarted = false;
let scrollBottomStarted = false;
let scrollTopInterval;
let scrollBottomInterval;
let scrollRightInterval;
let scrollLeftInterval;

let cellData = {
    "Sheet1": {}
};

let defaultProperties = {
    'font-family': 'Noto Sans',
    'font-size': 14,
    'text': "",
    'bold': false,
    'italic': false,
    'underline': false,
    'alignment': 'left',
    'color': "#444",
    'bgcolor': "#fff",
    'formula': "",
    'upStream': [],
    "downStream": []
};

//------------------ADDING ROWS, COLUMNS AND CELLS------------------//

for (let i = 1; i <= 100; i++) {
    let columnName = getColumnName(i);
    $("#columns").append(`<div class = 'column-name column-${i}' id='${columnName}'>${columnName}</div>`);
    $("#rows").append(`<div class='row-name'>${i}</div>`);
}

for (let i = 1; i <= 100; i++) {
    let row = $('<div class="cell-row"></div>');
    for (let j = 1; j <= 100; j++) {
        row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`);
    }
    $("#cells").append(row);
}

//------------------FUNCTION TO GENERATE COLUMN NAME------------------//

function getColumnName(columnNumber) {
    let str = "";
    while (columnNumber != 0) {
        columnNumber--;
        str = String.fromCharCode((columnNumber % 26) + 65) + str;
        columnNumber = Math.floor(columnNumber / 26);
    }
    return str;
}

//------------------CELLS EVENTS------------------//

$("#cells").scroll(function (e) {
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
})

$(".input-cell").dblclick(function (e) {
    $('.input-cell.selected').removeClass('selected top-selected bottom-selected left-selected right-selected');
    $(this).addClass('selected');
    $(this).attr('contenteditable', 'true');
    $(this).focus();
})

$(".input-cell").click(function (e) {
    let [rowId, colId] = getRowCol(this);
    let [topCell, bottomCell, leftCell, rightCell] = getTopLeftBottomRightCell(rowId, colId);
    if ($(this).hasClass('selected') && e.ctrlKey) {
        unselectCell(this, e, topCell, bottomCell, leftCell, rightCell);
    }
    else {
        selectCell(this, e, topCell, bottomCell, leftCell, rightCell);
    }
})

$(".input-cell").blur(function (e) {
    $(this).attr('contenteditable', 'false');
    let [rowId, colId] = getRowCol(this);
    if (cellData[selectedSheet][rowId - 1][colId - 1].formula != "") {
        updateStreams(this, []);
    }
    cellData[selectedSheet][rowId - 1][colId - 1].formula = "";
    updateCellData("text", $(this).text());
    let selfColCode = $(`.column-${colId}`).attr("id");
    evalFormula(selfColCode + rowId);
})

$('.input-cell').mousemove(function (e) {
    e.preventDefault();
    if (e.buttons == 1) {

        if (e.pageX > ($(window).width() - 10) && !scrollRightStarted) {
            scrollToRight();
        }
        else if (e.pageX < 10 && !scrollLeftStarted) {
            scrollToLeft();
        }
        else if (e.pageY > ($(window).height() - 10) && !scrollBottomStarted) {
            scrollToBottom();
        }
        else if (e.pageY < 10 && !scrollTopStarted) {
            scrollToTop();
        }

        if (!startCellSelected) {
            let [rowId, colId] = getRowCol(this);
            startCell = { 'rowId': rowId, 'colId': colId };
            selectAllBetweenCells(startCell, startCell);
            startCellSelected = true;
            $(".input-cell.selected").attr('contenteditable', 'false');
        }
    }
    else {
        startCellSelected = false;
    }
})

$('.input-cell').mouseenter(function (e) {
    if (e.buttons == 1) {

        if (e.pageX < ($(window).width() - 10) && scrollRightStarted) {
            clearInterval(scrollRightInterval);
            scrollRightStarted = false;
        }

        if (e.pageX > 10 && scrollLeftStarted) {
            clearInterval(scrollLeftInterval);
            scrollLeftStarted = false;
        }

        if (e.pageY > 10 && scrollTopStarted) {
            clearInterval(scrollTopInterval);
            scrollTopStarted = false;
        }

        if (e.pageY < ($(window).height() - 10) && scrollBottomStarted) {
            clearInterval(scrollBottomInterval);
            scrollBottomStarted = false;
        }

        let [rowId, colId] = getRowCol(this);
        endCell = { 'rowId': rowId, 'colId': colId };
        selectAllBetweenCells(startCell, endCell);
    }
})

//------------------FUNCTION TO GET ROW AND COLUMN NUMBER OF CELL------------------//

function getRowCol(ele) {
    let id = $(ele).attr('id');
    let idArr = id.split('-');
    let row = parseInt(idArr[1]);
    let col = parseInt(idArr[3]);
    return [row, col];
}

//------------------FUNCTION TO GET ADJACENT CELLS OF A CELL------------------//

function getTopLeftBottomRightCell(rowId, colId) {
    let topCell = $(`#row-${rowId - 1}-col-${colId}`);
    let bottomCell = $(`#row-${rowId + 1}-col-${colId}`);
    let rightCell = $(`#row-${rowId}-col-${colId + 1}`);
    let leftCell = $(`#row-${rowId}-col-${colId - 1}`);
    return [topCell, bottomCell, leftCell, rightCell];
}

//------------------FUNCTION TO SELECT A CELL------------------//

function selectCell(ele, e, topCell, bottomCell, leftCell, rightCell) {

    if (e.ctrlKey) {
        let topselected;
        if (topCell) {
            topselected = topCell.hasClass('selected');
        }
        let bottomselected;
        if (bottomCell) {
            bottomselected = bottomCell.hasClass('selected');
        }
        let leftselected;
        if (leftCell) {
            leftselected = leftCell.hasClass('selected');
        }
        let rightselected;
        if (rightCell) {
            rightselected = rightCell.hasClass('selected');
        }
        if (topselected) {
            $(ele).addClass('top-selected');
            topCell.addClass('bottom-selected');
        }
        if (bottomselected) {
            $(ele).addClass('bottom-selected');
            bottomCell.addClass('top-selected');
        }
        if (leftselected) {
            $(ele).addClass('left-selected');
            leftCell.addClass('right-selected');
        }
        if (rightselected) {
            $(ele).addClass('right-selected');
            rightCell.addClass('left-selected');
        }
    }
    else {
        $(".input-cell.selected").removeClass('selected top-selected bottom-selected left-selected right-selected');
    }
    $(ele).addClass('selected');
    changeHeader(getRowCol(ele));
}

//------------------FUNCTION TO UNSELECT A CELL------------------//

function unselectCell(ele, e, topCell, bottomCell, leftCell, rightCell) {
    if ($(ele).attr('contenteditable') == 'false') {
        if ($(ele).hasClass('top-selected')) {
            topCell.removeClass('bottom-selected');
        }
        if ($(ele).hasClass('bottom-selected')) {
            bottomCell.removeClass('top-selected');
        }
        if ($(ele).hasClass('left-selected')) {
            leftCell.removeClass('right-selected');
        }
        if ($(ele).hasClass('right-selected')) {
            rightCell.removeClass('left-selected');
        }
        $(ele).removeClass('selected top-selected botton-selected left-selected right-selected');
    }
}

//------------------FUNCTION TO SELECT CELLS WHILE DRAGGING------------------//

function selectAllBetweenCells(startCell, endCell) {
    $('.input-cell.selected').removeClass('selected top-selected bottom-selected left-selected right-selected');
    for (let i = Math.min(startCell.rowId, endCell.rowId); i <= Math.max(startCell.rowId, endCell.rowId); i++) {
        for (let j = Math.min(startCell.colId, endCell.colId); j <= Math.max(startCell.colId, endCell.colId); j++) {
            let [topCell, bottomCell, leftCell, rightCell] = getTopLeftBottomRightCell(i, j);
            selectCell($(`#row-${i}-col-${j}`)[0], { 'ctrlKey': true }, topCell, bottomCell, leftCell, rightCell);
        }
    }
}

//------------------FORMATTING BAR EVENTS------------------//

$(".alignment").click(function () {
    $(".alignment.selected").removeClass('selected');
    let temp = $(this).attr('data-type');
    $(this).addClass('selected');
    $(".input-cell.selected").css('text-align', temp);
    updateCellData("alignment", temp);
})

$("#bold").click(function (e) {
    setStyle(this, 'font-weight', 'bold', 'bold');
})

$("#italic").click(function (e) {
    setStyle(this, 'font-style', 'italic', 'italic');
})

$("#underline").click(function (e) {
    setStyle(this, 'text-decoration', 'underline', 'underline');
})

$('#fill-color').click(function () {
    setTimeout(() => {
        $(this).parent().click();
    }, 30);
})

$('#text-color').click(function () {
    setTimeout(() => {
        $(this).parent().click();
    }, 30);
})

$(".menu-selector").change(function (e) {
    let val = $(this).val();
    let key = $(this).attr('id');
    if (key == 'font-family') {
        $("#font-family").css(key, val);
    }
    if (!isNaN(val)) {
        val = parseInt(val);
    }
    $(".input-cell.selected").css(key, val);
    updateCellData(key, val);
})

$(".pick-color").colorPick({
    'initialColor': "#abcd",
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': true,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    'onColorSelected': function () {
        if (this.color != '#ABCD') {
            if ($(this.element.children()[1]).attr('id') == 'fill-color') {
                $('.input-cell.selected').css('background-color', this.color);
                $("#fill-color").css('border-bottom', `4px solid ${this.color}`);
                updateCellData('bgcolor', this.color);
            }
            if ($(this.element.children()[1]).attr('id') == 'text-color') {
                $('.input-cell.selected').css('color', this.color);
                $("#text-color").css('border-bottom', `4px solid ${this.color}`);
                updateCellData('color', this.color);
            }
        }
    }
});

//------------------FUNCTION TO CHANGE FORMATTING BAR ACCORDING TO SELECTED CELL------------------//

function changeHeader([rowId, colId]) {
    let data;
    if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
        data = cellData[selectedSheet][rowId - 1][colId - 1];
    }
    else {
        data = defaultProperties;
    }

    $('.alignment.selected').removeClass('selected');
    $(`.alignment[data-type = ${data.alignment}]`).addClass('selected');

    addRemoveSelectSetStyle(data, "bold");
    addRemoveSelectSetStyle(data, "italic");
    addRemoveSelectSetStyle(data, "underline");

    $("#fill-color").css('border-bottom', `4px solid ${data.bgcolor}`);
    $("#text-color").css('border-bottom', `4px solid ${data.color}`);
    $("#font-family").val(data["font-family"]);
    $("#font-size").val(data["font-size"]);
    $("#font-family").css('font-family', data['font-family']);
    $("#formula-input").text(data.formula);
}

//------------------FUNCTION TO SET/REMOVE STYLES------------------//

function setStyle(ele, key, value, property) {
    if ($(ele).hasClass('selected')) {
        $(ele).removeClass('selected');
        $('.input-cell.selected').css(key, '');
        updateCellData(property, false);
    }
    else {
        $(ele).addClass('selected');
        $('.input-cell.selected').css(key, value);
        updateCellData(property, true);
    }
}

function addRemoveSelectSetStyle(data, property) {
    if (data[property]) {
        $(`#${property}`).addClass('selected');
    }
    else {
        $(`#${property}`).removeClass('selected');
    }
}

//------------------FUNCTION TO UPDATE DATA OF ONLY CHANGED CELLS------------------//

function updateCellData(property, value) {

    let currentCellData = JSON.stringify(cellData);

    if (value != defaultProperties[property]) {
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = getRowCol(data);
            if (cellData[selectedSheet][rowId - 1] == undefined) {
                cellData[selectedSheet][rowId - 1] = {};
                cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [], "downStream": [] };
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
            }
            else {
                if (cellData[selectedSheet][rowId - 1][colId - 1] == undefined) {
                    cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [], "downStream": [] };
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                }
                else {
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                }
            }
        })
    }
    else {
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = getRowCol(data);
            if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                if (JSON.stringify(defaultProperties) == JSON.stringify(cellData[selectedSheet][rowId - 1][colId - 1])) {
                    delete cellData[selectedSheet][rowId - 1][colId - 1];
                    if (Object.keys(cellData[selectedSheet][rowId - 1]).length == 0) {
                        delete cellData[selectedSheet][rowId - 1];
                    }
                }
            }
        })
    }

    if (saved == true && JSON.stringify(cellData) != currentCellData) {
        saved = false;
    }

}

//------------------SHEET BAR EVENTS------------------//

$('.container').click(function (e) {
    $(".sheet-options-modal").remove();
})

$(".add-sheet").click(function (e) {
    lastAddedSheet++;
    totalSheets++;
    cellData[`Sheet${lastAddedSheet}`] = {};
    $(".sheet-tab.selected").removeClass('selected');
    $(".sheet-tab-container").append(`<div class = 'sheet-tab selected'>Sheet${lastAddedSheet}</div>`);
    selectSheet();
    addSheetEvents();
    saved = false;
})

$('.left-scroller,.right-scroller').click(function (e) {
    let keysArr = Object.keys(cellData);
    let selectedSheetIndex = keysArr.indexOf(selectedSheet);
    if (selectedSheetIndex != 0 && $(this).text() == 'arrow_left') {
        selectSheet($(".sheet-tab.selected").prev()[0]);
    }
    else if ($(this).text() == 'arrow_right' && selectedSheetIndex != (keysArr.length - 1)) {
        selectSheet($(".sheet-tab.selected").next()[0]);
    }
    $(".sheet-tab.selected")[0].scrollIntoView();
})

//------------------ADD EVENTS TO FIRST SHEET------------------//

addSheetEvents();

//------------------FUNCTION TO SELECT SHEET------------------//

function selectSheet(ele) {
    if (ele && !($(ele).hasClass('selected'))) {
        $(".sheet-tab.selected").removeClass('selected');
        $(ele).addClass('selected');
    }
    emptyPreviousSheet();
    selectedSheet = $(".sheet-tab.selected").text();
    loadCurrentSheet();
    $(".sheet-tab.selected")[0].scrollIntoView();
    $("#row-1-col-1").click();
}

//------------------FUNCTION TO ADD SHEET EVENTS (ALL)------------------//

function addSheetEvents() {
    $(".sheet-tab.selected").on("contextmenu", function (e) {
        e.preventDefault();
        selectSheet(this);
        $(".sheet-options-modal").remove();

        let modal = $('<div class="sheet-options-modal"><div class="option sheet-rename">Rename</div><div class="option sheet-delete">Delete</div></div>');
        modal.css('left', e.pageX);

        $(".container").append(modal);

        $(".sheet-rename").click(function (e) {
            let renameModal = $(`<div class="sheet-modal-parent">
            <div class="sheet-rename-modal">
                <div class="sheet-modal-title">Rename Sheet</div>
                <div class="sheet-modal-input-container">
                    <span class="sheet-modal-input-title">Rename Sheet To:</span>
                    <input type="text" class="sheet-modal-input">
                </div>
                <div class="sheet-modal-confirmation">
                    <div class="button yes-button">OK</div>
                    <div class="button no-button">Cancel</div>
                </div>
                </div>
            </div>`);

            $(".container").append(renameModal);
            $(".sheet-modal-input").focus();

            $(".no-button").click(function (e) {
                $(".sheet-modal-parent").remove();
            })
            $(".yes-button").click(function (e) {
                renameSheet();
            })
            $(".sheet-modal-input").keypress(function (e) {
                if (e.key == 'Enter') {
                    renameSheet();
                }
            })
        })

        $(".sheet-delete").click(function (e) {
            if (totalSheets > 1) {
                let sheetName = $(".sheet-tab.selected").text();
                let deleteModal = $(`
            <div class="sheet-modal-parent">
                <div class="sheet-delete-modal">
                    <div class="sheet-modal-title">${sheetName}</div>
                    <div class="sheet-modal-detail-container">
                        <span class="sheet-modal-detail-title">Are You Sure?</span>
                    </div>
                    <div class="sheet-modal-confirmation">
                        <div class="button yes-button">
                            <div class="material-icons delete-icon">delete</div>
                            Delete
                        </div>
                        <div class="button no-button">Cancel</div>
                    </div>
                </div>
            </div>`);
                $(".container").append(deleteModal);
                $(".no-button").click(function (e) {
                    $(".sheet-modal-parent").remove();
                })
                $(".yes-button").click(function (e) {
                    deleteSheet();
                    saved = false;
                })
            }
            else {
                alert('Not Possible');
            }
        });
    })

    $(".sheet-tab.selected").click(function (e) {
        selectSheet(this);
    })

}

//------------------FUNCTION TO EMPTY PREVIOUS SHEET CELLS------------------//

function emptyPreviousSheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for (let i of rowKeys) {
        let rowId = parseInt(i);
        let colKeys = Object.keys(data[rowId]);
        for (let j of colKeys) {
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);
            cell.text("");
            cell.css({
                'font-family': 'Noto Sans',
                'font-size': '14',
                'font-weight': '',
                'font-style': '',
                'text-decoration': '',
                'text-align': 'left',
                'color': "#444",
                'background-color': "#fff"
            })
        }
    }
}

//------------------FUNCTION TO LOAD NEW SHEET DATA------------------//

function loadCurrentSheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for (let i of rowKeys) {
        let rowId = parseInt(i);
        let colKeys = Object.keys(data[rowId]);
        for (let j of colKeys) {
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);
            cell.text(data[rowId][colId].text);
            cell.css({
                'font-family': data[rowId][colId]['font-family'],
                'font-size': data[rowId][colId]['font-size'],
                'font-weight': data[rowId][colId].bold ? "bold" : "",
                'font-style': data[rowId][colId].italic ? "italic" : "",
                'text-decoration': data[rowId][colId].underline ? "underline" : "",
                'text-align': data[rowId][colId].alignment,
                'color': data[rowId][colId]['color'],
                'background-color': data[rowId][colId]['bgcolor']
            })
        }
    }
}

//------------------FUNCTION TO RENAME SHEET------------------//

function renameSheet() {
    let newName = $(".sheet-modal-input").val();
    if (newName && !(Object.keys(cellData).includes(newName))) {
        saved = false;
        let newCellData = {};
        for (let i of Object.keys(cellData)) {
            if (i == selectedSheet) {
                newCellData[newName] = cellData[selectedSheet];
            }
            else {
                newCellData[i] = cellData[i];
            }
        }
        cellData = newCellData;
        selectedSheet = newName;
        $(".sheet-tab.selected").text(newName);
        $(".sheet-modal-parent").remove();
    }
    else {
        $(".rename-error").remove();
        $(".sheet-modal-input-container").append(`<div class='rename-error'> Sheet Name Not Valid / Already Exists </div>`);
    }
}

//------------------FUNCTION TO DELETE SHEET------------------//

function deleteSheet() {
    $(".sheet-modal-parent").remove();
    let sheetIndex = Object.keys(cellData).indexOf(selectedSheet);
    let currentSheet = $(".sheet-tab.selected");
    if (sheetIndex == 0) {
        selectSheet(currentSheet.next()[0]);
    }
    else {
        selectSheet(currentSheet.prev()[0]);
    }
    delete cellData[currentSheet.text()];
    currentSheet.remove();
    totalSheets--;
}

//------------------FILE BAR EVENTS (ALL)------------------//

$("#menu-file").click(function (e) {
    let modal = $(`
        <div class="file-modal">
            <div class="file-options-modal">
                <div class="close">
                    <div class="material-icons close-icon">arrow_circle_down</div>
                    <div>Close</div>
                </div>
                <div class="new">
                    <div class="material-icons new-icon">insert_drive_file</div>
                    <div>New</div>
                </div>
                <div class="open">
                    <div class="material-icons open-icon">folder_open</div>
                    <div>Open</div>
                </div>
                <div class="save">
                    <div class="material-icons save-icon">save</div>
                    <div>Save</div>
                </div>
            </div>
            <div class="file-recent-modal"></div>
            <div class="file-transparent"></div>
        </div>
    `);

    $('.container').append(modal);
    modal.animate({
        width: '100vw'
    }, 300)

    $(".close,.file-transparent,.new,.open").click(function (e) {
        modal.animate({
            width: '0vw'
        }, 300)
        setTimeout(() => {
            modal.remove();
        }, 250);
    })

    $(".new").click(function (e) {
        if (saved) {
            newFile();
        }
        else {
            $(`.container`).append(`
        <div class="sheet-modal-parent">
            <div class="sheet-delete-modal">
                <div class="sheet-modal-title">${$(".title").text()}</div>
                <div class="sheet-modal-detail-container">
                    <span class="sheet-modal-detail-title">Do you want to save changes?</span>
                </div>
                <div class="sheet-modal-confirmation">
                    <div class="button yes-button">
                        Yes
                    </div>
                    <div class="button no-button">No</div>
                </div>
            </div>
        </div>`)

            $(".no-button").click(function (e) {
                $(".sheet-modal-parent").remove();
                newFile();
            })

            $(".yes-button").click(function (e) {
                $(".sheet-modal-parent").remove();
                saveFile(true);
            })
        }
    })

    $(".save").click(function (e) {
        if (!saved) {
            saveFile();
        }
    })

    $(".open").click(function (e) {
        openFile();
    })

})

//------------------FUNCTION TO ADD NEW FILE------------------//

function newFile() {
    emptyPreviousSheet();
    cellData = { "Sheet1": {} };
    $(".sheet-tab").remove();
    selectedSheet = "Sheet1";
    totalSheets = 1;
    lastAddedSheet = 1;
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet1</div>`);
    $(".title").text("Excel - Book");
    $("#row-1-col-1").click();
    addSheetEvents();
}

//------------------FUNCTION TO SAVE FILE------------------//

function saveFile(newClicked) {
    $('.container').append(`
        <div class="sheet-modal-parent">
            <div class="sheet-rename-modal">
                <div class="sheet-modal-title">Save File</div>
                <div class="sheet-modal-input-container">
                    <span class="sheet-modal-input-title">File Name:</span>
                    <input type="text" class="sheet-modal-input" value="${$('.title').text()}">
                </div>
                <div class="sheet-modal-confirmation">
                    <div class="button yes-button">Save</div>
                    <div class="button no-button">Cancel</div>
                </div>
            </div>
        </div>
    `)

    $(".yes-button").click(function (e) {
        $(".title").text($(".sheet-modal-input").val());
        let a = document.createElement("a");
        a.href = `data:application/json,${encodeURIComponent(JSON.stringify(cellData))}`;
        a.download = $(".title").text() + ".json";
        $(".container").append(a);
        a.click();
        a.remove();
        saved = true;
    })

    $(".no-button,.yes-button").click(function (e) {
        $(".sheet-modal-parent").remove();
        if (newClicked) {
            newFile();
        }
    })
}

//------------------FUNCTION TO OPEN SAVED FILE------------------//

function openFile() {
    let inputFile = $(`<input accept="application/json" type='file'>`);
    $(".container").append(inputFile);
    inputFile.click();
    inputFile.change(function (e) {
        let file = e.target.files[0];
        $(".title").text(file.name.split(".json")[0]);
        let reader = new FileReader();
        reader.readAsText(file);
        reader.onload = () => {
            cellData = JSON.parse(reader.result);
            emptyPreviousSheet();
            $(".sheet-tab").remove();
            let keys = Object.keys(cellData);
            for (let i = 0; i < keys.length; i++) {
                if (keys[i].includes("Sheet")) {
                    let splittedSheet = keys[i].split("Sheet");
                    if (splittedSheet.length == 2 && !isNaN(splittedSheet[1])) {
                        lastAddedSheet = parseInt(splittedSheet[1]);
                    }
                }
                $(".sheet-tab-container").append(`<div class="sheet-tab selected">${keys[i]}</div>`);
            }
            addSheetEvents();
            $('.sheet-tab').removeClass('selected');
            $($(".sheet-tab")[0]).addClass('selected');
            selectedSheet = "Sheet1";
            totalSheets = keys.length;
            loadCurrentSheet();
            inputFile.remove();
        }
    })
}

//------------------CUT, COPY AND PASTE EVENTS------------------//

$("#cut").click(function (e) {
    cutClicked = true;
    createClipboardContents();
})

$("#copy").click(function (e) {
    cutClicked = false;
    createClipboardContents();
})

$("#paste").click(function (e) {
    pasteIntoSheet();
})

//------------------FUNCTION TO SAVE DATA COPIED/CUT------------------//

function createClipboardContents() {
    clipboard = { startCell: [], cellData: {} };
    clipboard.startCell = getRowCol($(".input-cell.selected")[0]);
    $(".input-cell.selected").each(function (index, data) {
        let [rowId, colId] = getRowCol(data);
        if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
            if (!clipboard.cellData[rowId]) {
                clipboard.cellData[rowId] = {};
            }
            clipboard.cellData[rowId][colId] = { ...cellData[selectedSheet][rowId - 1][colId - 1] };
        }
    })
    $(".input-cell.selected").removeClass('selected top-selected left-selected right-selected bottom-selected');
}

//------------------FUNCTION TO PASTE DATA COPIED/CUT------------------//

function pasteIntoSheet() {
    if (cutClicked) {
        emptyPreviousSheet();
    }
    let startCell = getRowCol($(".input-cell.selected")[0]);
    let rows = Object.keys(clipboard.cellData);
    for (i of rows) {
        let cols = Object.keys(clipboard.cellData[i]);
        for (j of cols) {
            if (cutClicked) {
                delete cellData[selectedSheet][i - 1][j - 1];
                if (Object.keys(cellData[selectedSheet][i - 1]).length == 0) {
                    delete cellData[selectedSheet][i - 1];
                }
            }
        }
    }
    for (i of rows) {
        let cols = Object.keys(clipboard.cellData[i]);
        for (j of cols) {
            let rowDistance = parseInt(i) - parseInt(clipboard.startCell[0]);
            let colDistance = parseInt(j) - parseInt(clipboard.startCell[1]);
            if (!cellData[selectedSheet][startCell[0] + rowDistance - 1]) {
                cellData[selectedSheet][startCell[0] + rowDistance - 1] = {};
            }
            cellData[selectedSheet][startCell[0] + rowDistance - 1][startCell[1] + colDistance - 1] = { ...clipboard.cellData[i][j] };
        }
    }
    loadCurrentSheet();
    if (cutClicked) {
        clipboard = { startCell: [], cellData: {} };
        cutClicked = false;
    }
}

//------------------FORMULA BAR EVENT------------------//

$("#formula-input").blur(function (e) {
    if ($(".input-cell.selected").length > 0) {
        let formula = $(this).text();
        let tempElements = formula.split(" ");
        let elements = [];
        for (let i of tempElements) {
            if (i.length >= 2) {
                i = i.replace("(", "");
                i = i.replace(")", "");
                if (!elements.includes(i)) {
                    elements.push(i);
                }
            }
        }
        $(".input-cell.selected").each(function (index, data) {
            if (updateStreams(data, elements, false)) {
                let [rowId, colId] = getRowCol(data);
                cellData[selectedSheet][rowId - 1][colId - 1].formula = formula;
                let selfColCode = $(`.column-${colId}`).attr("id");
                evalFormula(selfColCode + rowId);
            }
            else {
                alert("Wrong Formula.!");
            }
        })
    }
    else {
        alert("No Cell Selected.!");
    }
})

//------------------FUNCTION TO CHECK VALIDITY OF FORMULA------------------//

function updateStreams(ele, elements, update, oldUpstream) {
    let [rowId, colId] = getRowCol(ele);
    let selfColCode = $(`.column-${colId}`).attr('id');

    if (elements.includes(selfColCode + rowId)) {
        return false;
    }

    if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
        let ds = cellData[selectedSheet][rowId - 1][colId - 1].downStream;
        let us = cellData[selectedSheet][rowId - 1][colId - 1].upStream;
        for (let i of ds) {
            if (elements.includes(i)) {
                return false;
            }
        }
        for (let i of ds) {
            let [calRowId, calColId] = codeToValue(i);
            updateStreams($(`#row-${calRowId}-col-${calColId}`)[0], elements, true, us);
        }
    }

    if (!cellData[selectedSheet][rowId - 1]) {
        cellData[selectedSheet][rowId - 1] = {};
        cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [...elements], "downStream": [] };
    }
    else if (!cellData[selectedSheet][rowId - 1][colId - 1]) {
        cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [...elements], "downStream": [] };
    }
    else {
        let upStream = [...cellData[selectedSheet][rowId - 1][colId - 1].upStream];
        if (update) {
            for (let i of oldUpstream) {
                let [calRowId, calColId] = codeToValue(i);
                let index = cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.indexOf(selfColCode + rowId);
                cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);
                if (JSON.stringify(cellData[selectedSheet][calRowId - 1][calColId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][calRowId - 1][calColId - 1];
                    if (Object.keys(cellData[selectedSheet][calRowId - 1]).length == 0) {
                        delete cellData[selectedSheet][calRowId - 1];
                    }
                }
                index = cellData[selectedSheet][rowId - 1][colId - 1].upStream.indexOf(i);
                cellData[selectedSheet][rowId - 1][colId - 1].upStream.splice(index, 1);
            }
            for (let i of elements) {
                cellData[selectedSheet][rowId - 1][colId - 1].upStream.push(i);
            }
        }
        else {
            for (let i of upStream) {
                let [calRowId, calColId] = codeToValue(i);
                let index = cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.indexOf(selfColCode + rowId);
                cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);
                if (JSON.stringify(cellData[selectedSheet][calRowId - 1][calColId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][calRowId - 1][calColId - 1];
                    if (Object.keys(cellData[selectedSheet][calRowId - 1]).length == 0) {
                        delete cellData[selectedSheet][calRowId - 1];
                    }
                }
            }
            cellData[selectedSheet][rowId - 1][colId - 1].upStream = [...elements];
        }
    }

    for (let i of elements) {
        let [calRowId, calColId] = codeToValue(i);
        if (!cellData[selectedSheet][calRowId - 1]) {
            cellData[selectedSheet][calRowId - 1] = {};
            cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [selfColCode + rowId] };
        }
        else if (!cellData[selectedSheet][calRowId - 1][calColId - 1]) {
            cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [selfColCode + rowId] };
        }
        else {
            cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.push(selfColCode + rowId);
        }
    }
    return true;
}

//------------------FUNCTION TO CHANGE CELL NAME TO ROW AND COLUMN NUMBER------------------//

function codeToValue(code) {
    let colCode = "";
    let rowCode = "";
    for (let i = 0; i < code.length; i++) {
        if (!isNaN(code.charAt(i))) {
            rowCode += code.charAt(i);
        }
        else {
            colCode += code.charAt(i);
        }
    }
    let colId = parseInt($(`#${colCode}`).attr('class').split(' ')[1].split('-')[1]);
    let rowId = parseInt(rowCode);
    return [rowId, colId];
}

//------------------FUNCTION TO CHANGE DATA ACCORDING TO FORMULA------------------//

function evalFormula(cell) {
    let [rowId, colId] = codeToValue(cell);
    let formula = cellData[selectedSheet][rowId - 1][colId - 1].formula;
    if (formula != "") {
        let upStream = cellData[selectedSheet][rowId - 1][colId - 1].upStream;
        let upStreamValue = [];
        for (let i in upStream) {
            let [calRowId, calColId] = codeToValue(upStream[i]);
            let value;
            if (cellData[selectedSheet][calRowId - 1][calColId - 1].text == "") {
                value = "0";
            }
            else {
                value = cellData[selectedSheet][calRowId - 1][calColId - 1].text;
            }
            upStreamValue.push(value);
            formula = formula.replace(upStream[i], upStreamValue[i]);
        }
        cellData[selectedSheet][rowId - 1][colId - 1].text = eval(formula);
        loadCurrentSheet();
    }
    let downStream = cellData[selectedSheet][rowId - 1][colId - 1].downStream;
    for (let i = downStream.length - 1; i >= 0; i--) {
        evalFormula(downStream[i]);
    }
}

//------------------SCROLLING EVENTS------------------//

$(".data-container").mousemove(function (e) {
    e.preventDefault();
    if (e.buttons == 1) {
        if (e.pageX > ($(window).width() - 10) && !scrollRightStarted) {
            scrollToRight();
        } else if (e.pageX < (10) && !scrollLeftStarted) {
            scrollToLeft();
        }
        else if (e.pageY > ($(window).height() - 10) && !scrollBottomStarted) {
            scrollToBottom();
        }
        else if (e.pageY < 10 && !scrollTopStarted) {
            scrollToTop();
        }
    }
});

$(".data-container").mouseup(function (e) {
    clearInterval(scrollRightInterval);
    clearInterval(scrollLeftInterval);
    clearInterval(scrollTopInterval);
    clearInterval(scrollBottomInterval);
    scrollRightStarted = false;
    scrollLeftStarted = false;
    scrollTopStarted = false;
    scrollBottomStarted = false;
})

//------------------FUNCTIONS TO SCROLL WHEN SELECTING CELLS BY DRAGGING------------------//

function scrollToRight() {
    scrollRightStarted = true;
    scrollRightInterval = setInterval(function () {
        $("#cells").scrollLeft($("#cells").scrollLeft() + 100);
    }, 100)
}

function scrollToLeft() {
    scrollLeftStarted = true;
    scrollLeftInterval = setInterval(function () {
        $("#cells").scrollLeft($("#cells").scrollLeft() - 100);
    }, 100);

}

function scrollToTop() {
    scrollTopStarted = true;
    scrollTopInterval = setInterval(function () {
        $("#cells").scrollTop($("#cells").scrollTop() - 100);
    }, 100);
}

function scrollToBottom() {
    scrollBottomStarted = true;
    scrollBottomInterval = setInterval(function () {
        $("#cells").scrollTop($("#cells").scrollTop() + 100);
    }, 100);
}

//--------------------------------------------------------------------------------------//