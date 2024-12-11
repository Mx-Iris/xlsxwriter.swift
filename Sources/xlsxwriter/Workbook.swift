//
//  Workbook.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import libxlsxwriter

/// Struct to represent an Excel workbook.
public struct Workbook {
    var lxw_workbook: UnsafeMutablePointer<lxw_workbook>

    /// Create a new workbook object.
    public init?(filePath: String) {
        guard let workbook = filePath.withCString ({ workbook_new($0) }) else { return nil }
        self.lxw_workbook = workbook
    }
    
    /// Close the Workbook object and write the XLSX file.
    public func close() throws {
        try XLSXWriterError.throwIfNeeded {
            workbook_close(lxw_workbook)
        }
    }

    /// Add a new worksheet to the Excel workbook.
    public func addWorksheet(name: String? = nil) -> Worksheet? {
        var worksheet: UnsafeMutablePointer<lxw_worksheet>?
        if let name = name {
            worksheet = name.withCString { workbook_add_worksheet(lxw_workbook, $0) }
        } else {
            worksheet = workbook_add_worksheet(lxw_workbook, nil)
        }
        
        return worksheet.map { Worksheet($0) }
    }

    /// Add a new chartsheet to a workbook.
    public func addChartsheet(name: String? = nil) -> Chartsheet? {
        let chartsheet: UnsafeMutablePointer<lxw_chartsheet>?
        if let name = name {
            chartsheet = name.withCString { workbook_add_chartsheet(lxw_workbook, $0) }
        } else {
            chartsheet = workbook_add_chartsheet(lxw_workbook, nil)
        }
        return chartsheet.map { Chartsheet($0) }
    }

    /// Add a new format to the Excel workbook.
    public func addFormat() -> Format? {
        workbook_add_format(lxw_workbook).map { Format($0) }
    }
    
    /// Create a new chart to be added to a worksheet
    public func addChart(of type: ChartType) -> Chart? {
        workbook_add_chart(lxw_workbook, type.rawValue).map { Chart($0) }
    }

    /// Get a worksheet object from its name.
    public subscript(worksheet name: String) -> Worksheet? {
        guard let ws = name.withCString({ s in workbook_get_worksheet_by_name(lxw_workbook, s) }) else {
            return nil
        }
        return Worksheet(ws)
    }

    /// Get a chartsheet object from its name.
    public subscript(chartsheet name: String) -> Chartsheet? {
        guard let cs = name.withCString({ s in workbook_get_chartsheet_by_name(lxw_workbook, s) })
        else { return nil }
        return Chartsheet(cs)
    }

    /// Validate a worksheet or chartsheet name.
    func validate(sheetName: String) throws {
        try sheetName.withCString { sheetName in
            try XLSXWriterError.throwIfNeeded {
                workbook_validate_sheet_name(lxw_workbook, sheetName)
            }
        }
    }
}
