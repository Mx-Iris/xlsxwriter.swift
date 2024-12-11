//
//  Cell_Range.swift
//
//
//  Created by Daniel Müllenborn on 02.01.21.
//

import libxlsxwriter

public struct Cell: ExpressibleByStringLiteral, ExpressibleByArrayLiteral {
    public let row: UInt32
    
    public let column: UInt16
    
    public init(stringLiteral value: String) {
        (self.row, self.column) = value.withCString { (lxw_name_to_row($0), lxw_name_to_col($0)) }
    }

    public init(arrayLiteral elements: Int...) {
        precondition(elements.count == 2, "[row, column]")
        self.row = UInt32(elements[0])
        self.column = UInt16(elements[1])
    }

    public init(_ row: UInt32, _ col: UInt16) {
        self.row = row
        self.column = col
    }
    
    
    public static func cell(row: Int, column: Int) -> Cell {
        Cell(UInt32(row), UInt16(column))
    }
}

public struct ColumnRange: ExpressibleByStringLiteral, ExpressibleByArrayLiteral {
    public let startColumn: UInt16
    public let endColumn: UInt16
    public init(stringLiteral value: String) {
        (self.startColumn, self.endColumn) = value.withCString { (lxw_name_to_col($0), lxw_name_to_col_2($0)) }
    }

    public init(arrayLiteral elements: Int...) {
        precondition(elements.count == 2, "[startColumn, endColumn]")
        self.startColumn = UInt16(elements[0])
        self.endColumn = UInt16(elements[1])
    }

    public init(_ col: UInt16, _ col2: UInt16) {
        self.startColumn = col
        self.endColumn = col2
    }
    
    public static func columnRange(startColumn: Int, endColumn: Int) -> ColumnRange {
        ColumnRange(UInt16(startColumn), UInt16(endColumn))
    }
}

public struct CellRange: ExpressibleByStringLiteral, ExpressibleByArrayLiteral {
    public let startRow: UInt32
    public let startColumn: UInt16
    public let endRow: UInt32
    public let endColumn: UInt16
    
    public init(_ startRow: UInt32, _ startColumn: UInt16, _ endRow: UInt32, _ endColumn: UInt16) {
        self.startRow = startRow
        self.startColumn = startColumn
        self.endRow = endRow
        self.endColumn = endColumn
    }
    
    public init(stringLiteral value: String) {
        (self.startRow, self.startColumn, self.endRow, self.endColumn) = value.withCString {
            (lxw_name_to_row($0), lxw_name_to_col($0), lxw_name_to_row_2($0), lxw_name_to_col_2($0))
        }
    }

    public init(arrayLiteral elements: Int...) {
        precondition(elements.count == 4, "[startRow, startColumn, endRow, endColumn]")
        self.startRow = UInt32(elements[0])
        self.startColumn = UInt16(elements[1])
        self.endRow = UInt32(elements[2])
        self.endColumn = UInt16(elements[3])
    }
    
    public static func cellRange(startRow: Int, startColumn: Int, endRow: Int, endColumn: Int) -> CellRange {
        CellRange(UInt32(startRow), UInt16(startColumn), UInt32(endRow), UInt16(endColumn))
    }
}
