//
//  Cell_Range.swift
//
//
//  Created by Daniel Müllenborn on 02.01.21.
//

import libxlsxwriter

public struct Cell: ExpressibleByStringLiteral, ExpressibleByArrayLiteral {
    let row: UInt32
    let column: UInt16
    public init(stringLiteral value: String) {
        (self.row, self.column) = value.withCString { (lxw_name_to_row($0), lxw_name_to_col($0)) }
    }

    public init(arrayLiteral elements: Int...) {
        precondition(elements.count == 2, "[row, column]")
        self.row = UInt32(elements[0])
        self.column = UInt16(elements[1])
    }

    init(_ row: UInt32, _ col: UInt16) {
        self.row = row
        self.column = col
    }
}

public struct ColumnRange: ExpressibleByStringLiteral, ExpressibleByArrayLiteral {
    let startColumn: UInt16
    let endColumn: UInt16
    public init(stringLiteral value: String) {
        (self.startColumn, self.endColumn) = value.withCString { (lxw_name_to_col($0), lxw_name_to_col_2($0)) }
    }

    public init(arrayLiteral elements: Int...) {
        precondition(elements.count == 2, "[startColumn, endColumn]")
        self.startColumn = UInt16(elements[0])
        self.endColumn = UInt16(elements[1])
    }

    init(_ col: UInt16, _ col2: UInt16) {
        self.startColumn = col
        self.endColumn = col2
    }
}

public struct CellRange: ExpressibleByStringLiteral, ExpressibleByArrayLiteral {
    let startRow: UInt32
    let startColumn: UInt16
    let endRow: UInt32
    let endColumn: UInt16
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
}
