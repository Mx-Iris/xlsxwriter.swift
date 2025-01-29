//
//  Worksheet.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import libxlsxwriter

/// Struct to represent an Excel worksheet.
public struct Worksheet {
    private var lxw_worksheet: UnsafeMutablePointer<lxw_worksheet>

    var name: String {
        .init(cString: lxw_worksheet.pointee.name)
    }

    init(_ lxw_worksheet: UnsafeMutablePointer<lxw_worksheet>) {
        self.lxw_worksheet = lxw_worksheet
    }

    /// Insert a chart object into a worksheet.
    public func insertChart(_ chart: Chart, cell: Cell) throws -> Worksheet {
        let r = UInt32(cell.row)
        let c = UInt16(cell.column)
        try XLSXWriterError.throwIfNeeded {
            worksheet_insert_chart(lxw_worksheet, r, c, chart.lxw_chart)
        }
        return self
    }

    /// Insert a chart object into a worksheet, with options.
    public func insertChart(_ chart: Chart, cell: Cell, scale: (x: Double, y: Double)) throws -> Worksheet {
        let r = UInt32(cell.row)
        let c = UInt16(cell.column)
        var o = lxw_chart_options(
            x_offset: 0, y_offset: 0, x_scale: scale.x, y_scale: scale.y, object_position: 2,
            description: nil, decorative: 0
        )
        try XLSXWriterError.throwIfNeeded {
            worksheet_insert_chart_opt(lxw_worksheet, r, c, chart.lxw_chart, &o)
        }
        return self
    }

    @discardableResult
    public func writeNumber(_ number: Double, cell: Cell, format: Format? = nil) throws -> Worksheet {
        let f = format?.lxw_format
        let r = UInt32(cell.row)
        var c = UInt16(cell.column)
        try XLSXWriterError.throwIfNeeded {
            worksheet_write_number(lxw_worksheet, r, c, number, f)
        }
        return self
    }

    /// Write a row of String values starting from (row, col).
    @discardableResult
    public func writeString(_ string: String, cell: Cell, format: Format? = nil) throws -> Worksheet {
        let f = format?.lxw_format
        let r = UInt32(cell.row)
        var c = UInt16(cell.column)

        try string.withCString { s in
            try XLSXWriterError.throwIfNeeded {
                worksheet_write_string(lxw_worksheet, r, c, s, f)
            }
        }

        return self
    }

    /// Write data to a worksheet cell by calling the appropriate
    /// worksheet_write_*() method based on the type of data being passed.
    @discardableResult
    public func writeValue(_ value: Value, cell: Cell, format: Format? = nil) throws -> Worksheet {
        let r = cell.row
        let c = cell.column
        let f = format?.lxw_format
        let error: lxw_error
        switch value {
        case let .number(number):
            error = worksheet_write_number(lxw_worksheet, r, c, number, f)
        case let .string(string):
            error = string.withCString { s in worksheet_write_string(lxw_worksheet, r, c, s, f) }
        case let .url(url):
            error = url.absoluteString.withCString { s in worksheet_write_url(lxw_worksheet, r, c, s, f) }
        case .blank: error = worksheet_write_blank(lxw_worksheet, r, c, f)
        case let .comment(comment):
            error = comment.withCString { s in worksheet_write_comment(lxw_worksheet, r, c, s) }
        case let .boolean(boolean):
            error = worksheet_write_boolean(lxw_worksheet, r, c, Int32(boolean ? 1 : 0), f)
        case let .formula(formula):
            error = formula.withCString { s in worksheet_write_formula(lxw_worksheet, r, c, s, f) }
        case let .datetime(datetime):
            error = lxw_error(rawValue: 0)
            let num = (datetime.timeIntervalSince1970 / 86400) + 25569
            worksheet_write_number(lxw_worksheet, r, c, num, f)
        }

        try XLSXWriterError.throwIfNeeded {
            error
        }

        return self
    }

    /// Set a worksheet tab as selected.
    @discardableResult
    public func select() -> Worksheet {
        worksheet_select(lxw_worksheet)
        return self
    }

    /// Hide the current worksheet.
    @discardableResult
    public func hide() -> Worksheet {
        worksheet_hide(lxw_worksheet)
        return self
    }

    /// Make a worksheet the active, i.e., visible worksheet.
    @discardableResult
    public func activate() -> Worksheet {
        worksheet_activate(lxw_worksheet)
        return self
    }

    /// Hide zero values in worksheet cells.
    @discardableResult
    public func hideZero() -> Worksheet {
        worksheet_hide_zero(lxw_worksheet)
        return self
    }

    /// Set the paper type for printing.
    @discardableResult
    public func paper(ofType type: PaperType) -> Worksheet {
        worksheet_set_paper(lxw_worksheet, type.rawValue)
        return self
    }
    
    public enum Dimension {
        case `default`(Double)
        case pixels(Int)
    }

    @discardableResult
    public func column(_ col: Int, width: Dimension, format: Format? = nil, hidden: Bool = false) throws -> Worksheet {
        try column(.init(arrayLiteral: col, col), width: width, format: format, hidden: hidden)
    }

    /// Set the properties for one or more columns of cells.
    @discardableResult
    public func column(_ columnRange: ColumnRange, width: Dimension, format: Format? = nil, hidden: Bool = false) throws -> Worksheet {
        let first = columnRange.startColumn
        let last = columnRange.endColumn
        let f = format?.lxw_format
        var options = lxw_row_col_options()
        options.hidden = hidden.uint8
        try XLSXWriterError.throwIfNeeded {
            switch width {
            case .default(let width):
                worksheet_set_column_opt(lxw_worksheet, first, last, width, f, &options)
            case .pixels(let width):
                worksheet_set_column_pixels_opt(lxw_worksheet, first, last, width.uint32, f, &options)
            }
        }
        return self
    }

    /// Set the properties for a row of cells
    @discardableResult
    public func row(_ row: Int, height: Dimension, format: Format? = nil, hidden: Bool = false) throws -> Worksheet {
        let f = format?.lxw_format
        var options = lxw_row_col_options()
        options.hidden = hidden.uint8
        try XLSXWriterError.throwIfNeeded {
            switch height {
            case .default(let height):
                worksheet_set_row_opt(lxw_worksheet, row.uint32, height, f, &options)
            case .pixels(let height):
                worksheet_set_row_pixels_opt(lxw_worksheet, row.uint32, height.uint32, f, &options)
            }
        }
        return self
    }

    /// Set the properties for one or more columns of cells.
    @discardableResult
    public func hideColumn(_ col: Int, width: Dimension = .default(8.43)) throws -> Worksheet {
        let first = UInt16(col)
        let last = UInt16(col)
        try XLSXWriterError.throwIfNeeded {
            var o = lxw_row_col_options(hidden: 1, level: 0, collapsed: 0)
            switch width {
            case .default(let width):
                return worksheet_set_column_opt(lxw_worksheet, first, last, width, nil, &o)
            case .pixels(let width):
                return worksheet_set_column_pixels_opt(lxw_worksheet, first, last, width.uint32, nil, &o)
            }
        }
        return self
    }

    /// Set the color of the worksheet tab.
    @discardableResult
    public func tabColor(_ color: Color) -> Worksheet {
        worksheet_set_tab_color(lxw_worksheet, color.hex)
        return self
    }

    /// Set the default row properties.
    @discardableResult
    public func defaultRowHeight(_ height: Double, hideUnusedRows: Bool = true)
        -> Worksheet {
        let hide: UInt8 = hideUnusedRows ? 1 : 0
        worksheet_set_default_row(lxw_worksheet, height, hide)
        return self
    }

    /// Set the print area for a worksheet.
    @discardableResult
    public func printArea(forRange range: CellRange) throws -> Worksheet {
        try XLSXWriterError.throwIfNeeded {
            worksheet_print_area(lxw_worksheet, range.startRow, range.startColumn, range.endRow, range.endColumn)
        }
        return self
    }

    /// Set the autofilter area in the worksheet.
    @discardableResult
    public func autoFilter(forRange range: CellRange) throws -> Worksheet {
        try XLSXWriterError.throwIfNeeded {
            worksheet_autofilter(lxw_worksheet, range.startRow, range.startColumn, range.endRow, range.endColumn)
        }
        return self
    }

    /// Set the option to display or hide gridlines on the screen and the printed page.
    @discardableResult
    public func gridline(onScreen screen: Bool, print: Bool = false) -> Worksheet {
        worksheet_gridlines(lxw_worksheet, UInt8((print ? 2 : 0) + (screen ? 1 : 0)))
        return self
    }

    /// Set a table in the worksheet.
    @discardableResult
    public func table(
        forRange range: CellRange,
        name: String? = nil,
        header: [(String, Format?)] = []
    ) throws -> Worksheet {
        try table(
            forRange: range, name: name, header: header.map { $0.0 }, format: header.map { $0.1 },
            totalRow: []
        )
    }

    /// Merge a range of cells in the worksheet.
    @discardableResult
    public func mergeCell(forRange range: CellRange, string: String, format: Format? = nil) throws -> Worksheet {
        try XLSXWriterError.throwIfNeeded {
            worksheet_merge_range(lxw_worksheet, range.startRow, range.startColumn, range.endRow, range.endColumn, string, format?.lxw_format)
        }
        return self
    }

    /// Set a table in the worksheet.
    @discardableResult
    public func table(
        forRange range: CellRange, name: String? = nil, header: [String] = [], format: [Format?] = [],
        totalRow: [TotalFunction] = []
    ) throws -> Worksheet {
        var options = lxw_table_options()
        if let name = name { options.name = makeCString(from: name) }
        options.style_type = UInt8(LXW_TABLE_STYLE_TYPE_MEDIUM.rawValue)
        options.style_type_number = 7
        options.total_row = totalRow.isEmpty ? UInt8(LXW_FALSE.rawValue) : UInt8(LXW_TRUE.rawValue)
        var table_columns = [lxw_table_column]()
        let buffer = UnsafeMutableBufferPointer<UnsafeMutablePointer<lxw_table_column>?>.allocate(
            capacity: header.count + 1)
        defer { buffer.deallocate() }
        if !header.isEmpty {
            table_columns = Array(repeating: lxw_table_column(), count: header.count)
            for i in header.indices {
                table_columns[i].header = makeCString(from: header[i])
                if format.endIndex > i {
                    table_columns[i].header_format = format[i]?.lxw_format
                }
                if totalRow.endIndex > i {
                    table_columns[i].total_function = totalRow[i].rawValue
                }
                withUnsafeMutablePointer(to: &table_columns[i]) {
                    buffer.baseAddress?.advanced(by: i).pointee = $0
                }
            }
            options.columns = buffer.baseAddress
        }
        try XLSXWriterError.throwIfNeeded {
            worksheet_add_table(
                lxw_worksheet, range.startRow, range.startColumn, range.endRow + (totalRow.isEmpty ? 0 : 1), range.endColumn,
                &options
            )
        }
        if let _ = name { options.name.deallocate() }
        table_columns.forEach { $0.header.deallocate() }
        return self
    }
    
    @discardableResult
    public func freezePanes(firstRow: Int, firstColumn: Int) -> Self {
        worksheet_freeze_panes(lxw_worksheet, firstRow.uint32, firstColumn.uint16)
        return self
    }
    
    @discardableResult
    public func freezePanes(firstRow: Int, firstColumn: Int, topRow: Int, leftColumn: Int, isSplit: Bool) -> Self {
        worksheet_freeze_panes_opt(lxw_worksheet, firstRow.uint32, firstColumn.uint16, topRow.uint32, leftColumn.uint16, isSplit ? 1 : 0)
        return self
    }
}

private func makeCString(from str: String) -> UnsafePointer<CChar> {
    str.withCString { $0 }
}

extension BinaryInteger {
    var uint32: UInt32 {
        .init(self)
    }
    
    var uint16: UInt16 {
        .init(self)
    }
}

extension Bool {
    var uint8: UInt8 {
        self ? 1 : 0
    }
}
