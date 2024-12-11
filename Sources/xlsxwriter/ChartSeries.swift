//
//  Chart_Series.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import libxlsxwriter

/// Struct to represent an Excel chart.
public struct Chart {
    let lxw_chart: UnsafeMutablePointer<lxw_chart>
    init(_ lxw_chart: UnsafeMutablePointer<lxw_chart>) { self.lxw_chart = lxw_chart }
    /// Add a data series to a chart.
    @discardableResult public func addSeriesValues(_ values: String? = nil, name: String? = nil) -> Series {
        let series: Series
        if let values = values {
            series = values.withCString { Series(chart_add_series(lxw_chart, nil, $0)) }
        } else {
            series = Series(chart_add_series(lxw_chart, nil, nil))
        }
        if let name = name { series.name(name) }
        return series
    }

    /// Remove/hide one or more series in a chart legend (the series will still display on the chart).
    @discardableResult public func removeLegends(_ legends: Int...) -> Chart {
        var array = legends.map { Int16($0) }
        array.append(-1)
        chart_legend_delete_series(lxw_chart, &array)
        return self
    }

    /// Set the minimum and maximum value for the axis range.
    @discardableResult public func xAxisRange(_ xAxis: ClosedRange<Double>) -> Chart {
        chart_axis_set_min(lxw_chart.pointee.x_axis, xAxis.lowerBound)
        chart_axis_set_max(lxw_chart.pointee.x_axis, xAxis.upperBound)
        return self
    }

    /// Set the minimum and maximum value for the axis range.
    @discardableResult public func yAxisRange(_ yAxis: ClosedRange<Double>) -> Chart {
        chart_axis_set_min(lxw_chart.pointee.y_axis, yAxis.lowerBound)
        chart_axis_set_max(lxw_chart.pointee.y_axis, yAxis.upperBound)
        return self
    }

    /// Set the name caption of the x axis.
    @discardableResult public func xAxisName(_ name: String) -> Chart {
        name.withCString { chart_axis_set_name(lxw_chart.pointee.x_axis, $0) }
        return self
    }

    /// Set the name caption of the y axis.
    @discardableResult public func yAxisName(_ name: String) -> Chart {
        name.withCString { chart_axis_set_name(lxw_chart.pointee.y_axis, $0) }
        return self
    }

    /// Set the title of the chart.
    @discardableResult public func titleName(_ name: String) -> Chart {
        name.withCString { chart_title_set_name(lxw_chart, $0) }

        return self
    }

    /// Turn off/hide axis.
    @discardableResult public func hideAxis(_ axis: Axis) -> Chart {
        switch axis {
        case .x: chart_axis_off(lxw_chart.pointee.x_axis)
        case .y: chart_axis_off(lxw_chart.pointee.y_axis)
        }
        return self
    }

    /// Turn on/off the major gridlines for an axis.
    @discardableResult public func majorGridlinesVisible(visible: Bool = true, forAxis axis: Axis) -> Chart {
        let visible = UInt8(visible ? 1 : 0)
        switch axis {
        case .x: chart_axis_major_gridlines_set_visible(lxw_chart.pointee.x_axis, visible)
        case .y: chart_axis_major_gridlines_set_visible(lxw_chart.pointee.y_axis, visible)
        }
        return self
    }

    /// Turn on/off the minor gridlines for an axis.
    @discardableResult public func minorGridlinesVisible(_ visible: Bool = true, forAxis axis: Axis) -> Chart {
        let visible = UInt8(visible ? 1 : 0)
        switch axis {
        case .x: chart_axis_minor_gridlines_set_visible(lxw_chart.pointee.x_axis, visible)
        case .y: chart_axis_minor_gridlines_set_visible(lxw_chart.pointee.y_axis, visible)
        }
        return self
    }

    /// Set the increment of the major units in the axis
    @discardableResult public func majorUnit(_ unit: Double, forAxis axis: Axis) -> Chart {
        switch axis {
        case .x: chart_axis_set_major_unit(lxw_chart.pointee.x_axis, unit)
        case .y: chart_axis_set_major_unit(lxw_chart.pointee.y_axis, unit)
        }
        return self
    }

    /// Set the increment of the minor units in the axis.
    @discardableResult public func minorUnit(_ unit: Double, forAxis axis: Axis) -> Chart {
        switch axis {
        case .x: chart_axis_set_minor_unit(lxw_chart.pointee.x_axis, unit)
        case .y: chart_axis_set_minor_unit(lxw_chart.pointee.y_axis, unit)
        }
        return self
    }

    /// Turn off an automatic chart title.
    @discardableResult public func titleOffset() -> Chart {
        chart_title_off(lxw_chart)
        return self
    }

    /// Set the position of the chart legend.
    @discardableResult public func legendPosition(_ position: LegendPosition) -> Chart {
        chart_legend_set_position(lxw_chart, position.rawValue)
        chart_axis_off(lxw_chart.pointee.y_axis)
        return self
    }

    /// Set the Pie/Doughnut chart rotation.
    @discardableResult public func rotation(_ rotation: Int) -> Chart {
        chart_set_rotation(lxw_chart, UInt16(rotation))
        return self
    }

    /// Turn on a data table below the horizontal axis.
    @discardableResult public func tableStyle(_ style: Int) -> Chart {
        chart_set_table(lxw_chart)
        return self
    }

    /// Turn on a data table below the horizontal axis.
    @discardableResult public func style(_ style: Int) -> Chart {
        chart_set_style(lxw_chart, UInt8(style))
        return self
    }
}

/// Struct to represent an Excel chart data series.
public struct Series {
    let lxw_chart_series: UnsafeMutablePointer<lxw_chart_series>
    init(_ lxw_chart_series: UnsafeMutablePointer<lxw_chart_series>) {
        self.lxw_chart_series = lxw_chart_series
    }

    /// Set the name of a chart series range.
    @discardableResult public func name(_ name: String) -> Series {
        let _ = name.withCString { chart_series_set_name(lxw_chart_series, $0) }
        return self
    }

    /// The function is used to specify the the series marker:
    @discardableResult public func marker(_ marker: Int, size: Int) -> Series {
        chart_series_set_marker_type(lxw_chart_series, UInt8(marker))
        chart_series_set_marker_size(lxw_chart_series, UInt8(size))
        return self
    }

    /// Set a series "values" range using row and column values.
    @discardableResult public func values(sheet: Worksheet, range: Range) -> Series {
        let _ = sheet.name.withCString {
            chart_series_set_values(lxw_chart_series, $0, range.row, range.col, range.row2, range.col2)
        }
        return self
    }

    /// Set a series "categories" range using row and column values.
    @discardableResult public func categories(sheet: Worksheet, range: Range) -> Series {
        let _ = sheet.name.withCString {
            chart_series_set_categories(
                lxw_chart_series, $0, range.row, range.col, range.row2, range.col2
            )
        }
        return self
    }

    /// Smooth a line or scatter chart series.
    @discardableResult public func smooth(_ smooth: Bool) -> Series {
        let smooth: UInt8 = smooth ? 1 : 0
        chart_series_set_smooth(lxw_chart_series, smooth)
        return self
    }

    /// Set a series name formula using row and column values.
    @discardableResult public func name(sheet: Worksheet, cell: Cell) -> Series {
        let _ = sheet.name.withCString {
            chart_series_set_name_range(lxw_chart_series, $0, cell.row, cell.col)
        }
        return self
    }

    /// Turn on a trendline for a chart data series.
    @discardableResult public func trendline(of type: TrendlineType, value: Int = 2) -> Series {
        chart_series_set_trendline(lxw_chart_series, type.rawValue, UInt8(value))
        return self
    }

    /// Set the trendline name for a chart data series.
    @discardableResult public func trendlineName(_ name: String) -> Series {
        let _ = name.withCString { chart_series_set_trendline_name(lxw_chart_series, $0) }
        return self
    }

    /// Set the trendline line properties for a chart data series.
    @discardableResult public func trendlineColor(
        _ color: Color = .black, width: Float = 2.25, dash_type: Int, transparency: Int = 0,
        hide: Bool = false
    )
        -> Series {
        var line = lxw_chart_line(
            color: Color.black.hex, none: hide ? 1 : 0, width: width, dash_type: UInt8(dash_type),
            transparency: UInt8(transparency)
        )
        chart_series_set_trendline_line(lxw_chart_series, &line)
        return self
    }

    /// Display the equation of a trendline for a chart data series.
    @discardableResult public func trendlineEquation() -> Series {
        chart_series_set_trendline_equation(lxw_chart_series)
        return self
    }

    /// Display the R squared value of a trendline for a chart data series.
    @discardableResult public func trendlineRSquared() -> Series {
        chart_series_set_trendline_r_squared(lxw_chart_series)
        return self
    }
}
