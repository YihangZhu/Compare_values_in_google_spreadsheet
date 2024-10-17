function onInstall(e) {
    onOpen(e);
}

var ui = SpreadsheetApp.getUi()
var colors = ['#F5F5F5', '#DCDCDC	', '#C8C8C8	', '#B8B8B8', '#A8A8A8', '#888888', '#707070', '#505050']

function onOpen(e) {
    ui.createAddonMenu()
        .addItem('Highlight higher values', 'highlight_max')
        .addItem('Color values in order', 'colour_cells_in_order')
        .addToUi()
}

function highlight_cell(cell) {
    // cell.setFontWeight("bold")
    cell.setBackground("#A8A8A8")
}

function colour_cells_in_order_per_vector(values, cells) {
    const result = Array.from(values.keys()).sort((a, b) => values[a] - values[b])  // from small to large
    console.log(result)

    for (var i = 0; i < values.length; i++) {
        var idx = result[i]
        var cell = cells[idx]
        console.log(cell.getValues())
        cell.setBackground(colors[i])
    }
}

function colour_cells_in_order() {
    var range = SpreadsheetApp.getActiveSheet().getActiveRange();
    var num_rows = range.getNumRows()
    var num_cols = range.getNumColumns()
    var result = ui.alert("Click \"Yes\", if dealing the values per column, \"No\" for per row", ui.ButtonSet.YES_NO_CANCEL)
    if (result == 'CANCEL') {
        return
    }
    var values = range.getValues()
    if (result == 'YES') {
        for (var c = 0; c < num_cols; c++) {
            var values_col = []
            var cells = []
            for (var r = 0; r < num_rows; r++) {
                values_col.push(values[r][c])
                cell = range.getCell(r + 1, c + 1)
                cells.push(cell)
            }
            colour_cells_in_order_per_vector(values_col, cells)
        }
    } else {
        for (var r = 0; r < num_rows; r++) {
            var values_row = []
            var cells = []
            for (var c = 0; c < num_cols; c++) {
                values_row.push(values[r][c])
                cell = range.getCell(r + 1, c + 1)
                cells.push(cell)
            }
            colour_cells_in_order_per_vector(values_row, cells)
        }
    }
}


function highlight_max() {
    var ranges = get_ranges()
    if (ranges == 0) {
        return 0
    }
    for (var r = 0; r < ranges.num_rows; r++) {
        for (var c = 0; c < ranges.num_cols; c++) {
            var max_value = Number.NEGATIVE_INFINITY
            var max_range = null
            for (var range_id = 0; range_id < ranges.num_ranges; range_id++) {
                console.log(ranges.values[range_id][r][c])
                if (ranges.values[range_id][r][c] > max_value) {
                    max_value = ranges.values[range_id][r][c]
                    max_range = ranges.ranges[range_id]
                }
            }
            cell = max_range.getCell(r + 1, c + 1)
            highlight_cell(cell)
        }
    }
    ui.alert("Comparison finished.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function get_ranges() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var result = ui.alert("Have you selected the ranges for comparison?", ui.ButtonSet.YES_NO_CANCEL);
    if (result == ui.Button.YES) {
        var ranges = sheet.getActiveRangeList().getRanges()
    } else {
        if (result == ui.Button.NO) {
            result = ui.prompt("Please  ranges via the notations: (an example of the input a1:c4,a5:c4,a10:c4)",
                ui.ButtonSet.OK_CANCEL);

            if (result.getSelectedButton() == ui.Button.CANCEL) {
                return 0;
            }
            result = String(result.getResponseText()).split(',')
            var ranges = []
            console.log(result)
            for (var range_id = 0; range_id < result.length; range_id++) {
                ranges.push(sheet.getRange(result[range_id]))
            }

        } else {
            if (result == ui.Button.CANCEL) {
                return 0;
            }
        }
    }

    var rows = ranges[0].getNumRows();
    var cols = ranges[0].getNumColumns();

    var ranges_values = []
    for (var range_id = 0; range_id < ranges.length; range_id++) {
        if (ranges[range_id].getNumRows() != rows) {
            result = ui.alert("The selected ranges have different number of rows", ui.ButtonSet.OK);
            return 0
        }
        if (ranges[range_id].getNumColumns() != cols) {
            result = ui.alert("The selected ranges have different number of columns", ui.ButtonSet.OK);
            return 0
        }
        ranges_values.push(ranges[range_id].getValues())
    }

    return {'ranges': ranges, 'values': ranges_values, 'num_rows': rows, 'num_cols': cols, 'num_ranges': ranges.length}
}



