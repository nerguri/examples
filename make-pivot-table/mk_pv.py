#!/usr/bin/env python

# -*- coding: utf-8 -*-

import sys
import argparse
import os
import traceback
import win32com.client as win32

win32c = win32.constants


DEF_VISIBLE = 0
DEF_PT_ROW_OFFSET = 0
DEF_PT_COL_OFFSET = 3
DEF_PC_ROW_OFFSET = 10
DEF_PC_COL_OFFSET = 10
DEF_PC_WIDTH = 1000
DEF_PC_HEIGHT = 400
DEF_PC_TYPE = 'xlLine'
DEF_PC_Y_AXIS_MAX = 100000

arg_parser = argparse.ArgumentParser(description='making pivot tables & charts')
arg_parser.add_argument('input_file_path', type=str,
                        help='input file path')
arg_parser.add_argument('sheet_indices', type=int,
                        nargs='+',
                        help='target sheet indices')
arg_parser.add_argument('-v', '--verbosity', action='count',
                        help='increase output verbosity')
arg_parser.add_argument('--visible', action='store_true',
                        help='setting Excel visible')
arg_parser.add_argument('-s', '--suffix', type=str,
                        default='_pv',
                        help='setting save file suffix')
arg_parser.add_argument('--pt_ro', type=int,
                        default=DEF_PT_ROW_OFFSET,
                        help='setting pivot table row offset')
arg_parser.add_argument('--pt_co', type=int,
                        default=DEF_PT_COL_OFFSET,
                        help='setting pivot table column offset')
arg_parser.add_argument('--pc_ro', type=int,
                        default=DEF_PC_ROW_OFFSET,
                        help='setting pivot chart row offset')
arg_parser.add_argument('--pc_co', type=int,
                        default=DEF_PC_COL_OFFSET,
                        help='setting pivot chart column offset')
arg_parser.add_argument('--pc_type', type=str,
                        default=DEF_PC_TYPE,
                        help='setting pivot chart type')
arg_parser.add_argument('--pc_width', type=int,
                        default=DEF_PC_WIDTH,
                        help='setting pivot chart width')
arg_parser.add_argument('--pc_height', type=int,
                        default=DEF_PC_HEIGHT,
                        help='setting pivot chart width')
arg_parser.add_argument('--pc_y_max', type=int,
                        default=DEF_PC_Y_AXIS_MAX,
                        help='setting pivot chart y axis max')
args = arg_parser.parse_args()

try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    if args.visible:
        excel.Visible = True

    abs_input_file_path = os.path.abspath(args.input_file_path)
    excel.Workbooks.Open(abs_input_file_path)

    wb = excel.ActiveWorkbook

    PC_TYPES_DICT = {
        'xlLine' : win32c.xlLine,
        'xlAreaStacked' : win32c.xlAreaStacked,
    }

    for index in args.sheet_indices:
        ws = wb.Sheets(index)
        sht_name = ws.Name
        ws.Select()
        ws.Cells(1, 1).Select()
    
        excel.Selection.CurrentRegion.Select()
    
        row_count = excel.Selection.Rows.Count
        col_count = excel.Selection.Columns.Count
    
    
        src_data = "'%s'!R%sC%s:R%sC%s" % (sht_name, 1, 1, row_count, col_count)
        pc = wb.PivotCaches().Create(SourceType = win32c.xlDatabase,
                                     SourceData = src_data,
                                     Version = win32c.xlPivotTableVersion14)
    
        pvt_tbl_name = '%s Pivot Table' % (sht_name, )
        pivot_x_pos = 1 + args.pt_ro
        pivot_y_pos = col_count + args.pt_co
        ws.Cells(pivot_x_pos, pivot_y_pos).Select()
        pt = pc.CreatePivotTable(TableDestination = "'%s'!R%sC%s" % (sht_name, pivot_x_pos, pivot_y_pos),
                                 TableName = pvt_tbl_name,
                                 DefaultVersion = win32c.xlPivotTableVersion14)
    
        field_name = 'Datetime(ms)'
        pf = pt.PivotFields(field_name)
        pf.Orientation = win32c.xlRowField
        pf.Position = 1
    
        field_name = 'Endpoint'
        pf = pt.PivotFields(field_name)
        pf.Orientation = win32c.xlColumnField
        pf.Position = 1
    
        field_name = 'Device'
        pf = pt.PivotFields(field_name)
        pf.Orientation = win32c.xlColumnField
        pf.Position = 2
    
        field_name = 'rkB/s'
        caption_name = 'sum : %s' % (field_name, )
        pt.AddDataField(pt.PivotFields(field_name),
                        caption_name,
                        win32c.xlSum)
    
        field_name = 'wkB/s'
        caption_name = 'sum : %s' % (field_name, )
        pt.AddDataField(pt.PivotFields(field_name),
                        caption_name,
                        win32c.xlSum)
    
        ws.Cells(pivot_x_pos, pivot_y_pos).Select()
    
        excel.Selection.CurrentRegion.Select()
        row_count = excel.Selection.Rows.Count
        col_count = excel.Selection.Columns.Count
        src_data = "'%s'!%s:%s" % (sht_name,
                                   ws.Cells(pivot_x_pos, pivot_y_pos).Address,
                                   ws.Cells(row_count, col_count).Address)
    
        shape = ws.Shapes.AddChart2(-1, PC_TYPES_DICT[args.pc_type],
                                    args.pc_ro, args.pc_co,
                                    args.pc_width + args.pc_ro, args.pc_height + args.pc_co)
        shape.Select()
        chart = shape.Chart
        chart.SetSourceData(Source = ws.Range(src_data))
        chart.HasTitle = True
        chart.ChartTitle.Text = 'iostat'
    
        y_axis = chart.Axes(win32c.xlValue)
        y_axis.HasTitle = True
        y_axis.AxisTitle.Text = 'Bandwidht (kB/s)'
        y_axis.MinimumScale = 0
        y_axis.MaximumScale = args.pc_y_max
        y_axis.TickLabels.NumberFormat = '#,##0'
    
        x_axis = chart.Axes(win32c.xlCategory)
        x_axis.HasTitle = True
        x_axis.AxisTitle.Text = 'Epochtime (ms)'
    
        excel.ActiveWindow.LargeScroll(ToRight=-1)
        ws.Cells(1, 1).Select()

    ws = wb.Sheets(1)
    ws.Select()
    ws.Cells(1, 1).Select()
    excel.DisplayAlerts = False
    filename_prefix, extension = os.path.splitext(abs_input_file_path)
    new_file_path = '%s%s%s' % (filename_prefix, args.suffix, extension)
    wb.SaveAs(new_file_path, ConflictResolution=win32c.xlLocalSessionChanges)

except Exception as ex:
    print traceback.print_exc()
    print ex
    print ex.strerror
    print ex[2][2].encode('cp949', 'ignore')
finally:
    excel.Application.Quit()

