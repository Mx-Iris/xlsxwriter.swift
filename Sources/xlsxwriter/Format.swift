//
//  Format.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import libxlsxwriter

/// Struct to represent the formatting properties of an Excel format.
public struct Format {
    let lxw_format: UnsafeMutablePointer<lxw_format>
    init(_ lxw_format: UnsafeMutablePointer<lxw_format>) { self.lxw_format = lxw_format }
    /// Turn on bold for the format font.
    @discardableResult public func bold() -> Format {
        format_set_bold(lxw_format)
        return self
    }

    /// Turn on italic for the format font.
    @discardableResult public func italic() -> Format {
        format_set_italic(lxw_format)
        return self
    }

    /// Set the cell border style.
    @discardableResult public func borderStyle(_ style: Border) -> Format {
        format_set_border(lxw_format, style.rawValue)
        return self
    }

    /// Set the cell top border style.
    @discardableResult public func topStyle(_ style: Border) -> Format {
        format_set_top(lxw_format, style.rawValue)
        return self
    }

    /// Set the cell bottom border style.
    @discardableResult public func bottomStyle(_ style: Border) -> Format {
        format_set_bottom(lxw_format, style.rawValue)
        return self
    }

    /// Set the cell left border style.
    @discardableResult public func leftStyle(_ style: Border) -> Format {
        format_set_left(lxw_format, style.rawValue)
        return self
    }

    /// Set the cell right border style.
    @discardableResult public func rightStyle(_ style: Border) -> Format {
        format_set_right(lxw_format, style.rawValue)
        return self
    }

    /// Set the horizontal alignment for data in the cell.
    @discardableResult public func alignHorizontal(_ horizontal: HorizontalAlignment) -> Format {
        format_set_align(lxw_format, horizontal.rawValue)
        return self
    }

    /// Set the vertical alignment for data in the cell.
    @discardableResult public func alignVertical(_ vertical: VerticalAlignment) -> Format {
        format_set_align(lxw_format, vertical.rawValue)
        return self
    }

    /// Set the vertical alignment and horizontal alignment to center.
    @discardableResult public func center() -> Format {
        format_set_align(lxw_format, HorizontalAlignment.center.rawValue)
        format_set_align(lxw_format, VerticalAlignment.center.rawValue)
        return self
    }

    /// Set the number format for a cell.
    @discardableResult public func numberFormat(_ numberFormat: String) -> Format {
        numberFormat.withCString { format_set_num_format(lxw_format, $0) }
        return self
    }

    /// Set the Excel built-in number format for a cell.
    @discardableResult public func numberFormat(_ index: Int) -> Format {
        format_set_num_format_index(lxw_format, UInt8(index))
        return self
    }

    /// Turn on the text "shrink to fit" for a cell.
    @discardableResult public func shrink() -> Format {
        format_set_shrink(lxw_format)
        return self
    }

    /// Set the font used in the cell.
    @discardableResult public func fontName(_ name: String) -> Format {
        name.withCString { format_set_font_name(lxw_format, $0) }
        return self
    }

    /// Set the color of the cell border.
    @discardableResult public func borderColor(_ color: Color) -> Format {
        format_set_border_color(lxw_format, color.hex)
        return self
    }

    /// Set the color of the font used in the cell.
    @discardableResult public func fontColor(_ color: Color) -> Format {
        format_set_font_color(lxw_format, color.hex)
        return self
    }

    /// Set the pattern background color for a cell.
    @discardableResult public func backgroundColor(_ color: Color) -> Format {
        format_set_pattern(lxw_format, 1)
        format_set_bg_color(lxw_format, color.hex)
        return self
    }

    /// Set the rotation of the text in a cell.
    @discardableResult public func rotationAngle(_ angle: Int) -> Format {
        format_set_rotation(lxw_format, Int16(angle))
        return self
    }

    /// Set the size of the font used in the cell.
    @discardableResult public func fontSize(_ size: Double) -> Format {
        format_set_font_size(lxw_format, size)
        return self
    }

    /// Set the text wrap to cell. This is required which cell's text contains line break to show correctly.
    @discardableResult public func setTextWrap() -> Format {
        format_set_text_wrap(lxw_format)
        return self
    }
}

/// Structure for color which contains common colors.
public struct Color {
    public var hex: UInt32
    public init(hex: UInt32) {
        self.hex = hex
    }

    public static var black: Self = .init(hex: 0x1000000)
    public static var blue: Self = .init(hex: 0x0000FF)
    public static var brown: Self = .init(hex: 0x800000)
    public static var cyan: Self = .init(hex: 0x00FFFF)
    public static var gray: Self = .init(hex: 0x808080)
    public static var green: Self = .init(hex: 0x008000)
    public static var lime: Self = .init(hex: 0x00FF00)
    public static var magenta: Self = .init(hex: 0xFF00FF)
    public static var navy: Self = .init(hex: 0x000080)
    public static var orange: Self = .init(hex: 0xFF6600)
    public static var purple: Self = .init(hex: 0x800080)
    public static var red: Self = .init(hex: 0xFF0000)
    public static var silver: Self = .init(hex: 0xC0C0C0)
    public static var white: Self = .init(hex: 0xFFFFFF)
    public static var yellow: Self = .init(hex: 0xFFFF00)
}
