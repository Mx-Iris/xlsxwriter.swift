//
//  Enumerations.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import Foundation

public enum Value: ExpressibleByFloatLiteral, ExpressibleByStringLiteral {
    case url(URL)
    case blank
    case comment(String)
    case number(Double)
    case string(String)
    case boolean(Bool)
    case formula(String)
    case datetime(Date)
    public init(floatLiteral value: Double) { self = .number(value) }
    public init(stringLiteral value: String) { self = .string(value) }
}

public enum Axis: CaseIterable {
    case x
    case y
}

public enum TrendlineType: UInt8, CaseIterable {
    case linear
    case log
    case poly
    case power
    case exp
    case average
}

/// Cell border styles for use with format.set_border()
public enum Border: UInt8, CaseIterable {
    case noBorder
    case thin
    case medium
    case dashed
    case dotted
    case thick
    case double
    case hair
    case mediumDashed
    case dashDot
    case mediumDashDot
    case dashDotDot
    case mediumDashDotDot
    case slantDashDot
}

/// Alignment values for format.set(alignment:)
public enum HorizontalAlignment: UInt8, CaseIterable {
    case none = 0
    case left
    case center
    case right
    case fill
    case justify
    case centerAcross
    case distributed
}

/// Alignment values for format.set(alignment:)
public enum VerticalAlignment: UInt8, CaseIterable {
    case top = 8
    case bottom
    case center
    case justify
    case distributed
}

/// The Excel paper format type.
public enum PaperType: UInt8, CaseIterable {
    case printerDefault = 0 // Printer default
    case letter // 8 1/2 x 11 in
    case letterSmall // 8 1/2 x 11 in
    case tabloid // 11 x 17 in
    case ledger // 17 x 11 in
    case legal // 8 1/2 x 14 in
    case statement // 5 1/2 x 8 1/2 in
    case executive // 7 1/4 x 10 1/2 in
    case a3 // 297 x 420 mm
    case a4 // 210 x 297 mm
    case a4Small // 210 x 297 mm
    case a5 // 148 x 210 mm
    case b4 // 250 x 354 mm
    case b5 // 182 x 257 mm
    case folio // 8 1/2 x 13 in
    case quarto // 215 x 275 mm
    case unnamed1 // 10x14 in
    case unnamed2 // 11x17 in
    case note // 8 1/2 x 11 in
    case envelope9 // 3 7/8 x 8 7/8
    case envelope10 // 4 1/8 x 9 1/2
    case envelope11 // 4 1/2 x 10 3/8
    case envelope12 // 4 3/4 x 11
    case envelope14 // 5 x 11 1/2
    case cSizeSheet // ---
    case dSizeSheet // ---
    case eSizeSheet // ---
    case envelopeDL // 110 x 220 mm
    case envelopeC3 // 324 x 458 mm
    case envelopeC4 // 229 x 324 mm
    case envelopeC5 // 162 x 229 mm
    case envelopeC6 // 114 x 162 mm
    case envelopeC65 // 114 x 229 mm
    case envelopeB4 // 250 x 353 mm
    case envelopeB5 // 176 x 250 mm
    case envelopeB6 // 176 x 125 mm
    case envelope // 110 x 230 mm
    case monarch // 3.875 x 7.5 in
    case envelopeInch // 3 5/8 x 6 1/2 in
    case fanfold // 14 7/8 x 11 in
    case germanStdFanfold // 8 1/2 x 12 in
    case germanLegalFanfold // 8 1/2 x 13 in
}

/// Available chart types.
public enum ChartType: UInt8, CaseIterable {
    case none
    case area
    case area_stacked
    case area_percentage_stacked
    case bar
    case bar_stacked
    case bar_percentage_stacked
    case column
    case column_stacked
    case column_percentage_stacked
    case doughnut
    case line
    case line_stacked
    case line_percentage_stacked
    case pie
    case scatter
    case scatter_straight
    case scatter_straight_with_markers
    case scatter_smooth
    case scatter_smooth_with_markers
    case radar
    case radar_with_markers
    case radar_filled
}

public enum LegendPosition: UInt8, CaseIterable {
    case none = 0
    case right
    case left
    case top
    case bottom
    case topRight
    case overlayRight
    case overlayLeft
    case overlayTopRight
}

public enum TotalFunction: UInt8, ExpressibleByIntegerLiteral, CaseIterable {
    case none = 0
    /// Use the average function as the table total.
    case average = 101
    /// Use the count numbers function as the table total.
    case nums = 102
    /// Use the count function as the table total.
    case count = 103
    /// Use the max function as the table total.
    case max = 104
    /// Use the min function as the table total.
    case min = 105
    /// Use the standard deviation function as the table total.
    case stdDev = 107
    /// Use the sum function as the table total.
    case sum = 109

    public init(integerLiteral value: Int) {
        if let function = TotalFunction(rawValue: UInt8(value)) {
            self = function
        } else {
            self = .none
        }
    }
}
