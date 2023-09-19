# -*- coding: utf-8 -*-
"""
Created on Thu Aug  3 20:33:19 2023

@author: dkane
"""

import os

import xlsxwriter


def get_xy_ranges(die_info):
    x_max = max([x for (x,y) in die_info])
    y_max = max([y for (x,y) in die_info])
    x_min = min([x for (x,y) in die_info])
    y_min = min([y for (x,y) in die_info])
    x_range = range(x_min, x_max+1)
    y_range = range(y_min, y_max+1)
    # print('x_range:', x_range, 'y_range:', y_range)
    return x_range, y_range

# def write_wafermap_border(die_info, ws):
#     x_max, y_max = get_max_xy(die_info)
#     row_max_x, col_max_y = [0] * y_max, [0] * x_max
#     row_min_x, col_min_y = [9999] * y_max, [9999] * x_max
#     for x, y in die_info:
#         row_max_x[y-1] = x if x > row_max_x[y-1] else row_max_x[y-1]
#         row_min_x[y-1] = x if x < row_min_x[y-1] else row_min_x[y-1]
#         col_max_y[x-1] = y if y > col_max_y[x-1] else col_max_y[x-1]
#         col_min_y[x-1] = y if y < col_min_y[x-1] else col_min_y[x-1]
#     for x, y in die_info:
#         if row_max_x[y-1] == x:
#             die_info[(x,y)]["format"].set_right(1)
#         if row_min_x[y-1] == x:
#             die_info[(x,y)]["format"].set_left(1)
#         if col_max_y[x-1] == y:
#             die_info[(x,y)]["format"].set_bottom(1)
#         if col_min_y[x-1] == y:
#             die_info[(x,y)]["format"].set_top(1)

def wafer_map(die_info, xlsx_fp, bin_opt='SW', open_xlsx=True, good_bin=1, top_is_y_min=True):
    '''
    Parameters
    ----------
    die_info : dict
        Dictionary with key-value pairs of the form:
            (x <int>, y <int>) : {'sbin_num' : <int>, 'sbin_name' : <string>, 'hbin_num' : <int>, 'hbin_name' : <string>}
    xlsx_fp : string
        file path to store xlsx wafer map
    bin_opt : string, optional
        whether to use software ('SW') or hardware ('HW') bins for wafer map.
    top_is_y_min: bool
        whether top y coordinate in excel wafer map is min y (incrementing down the sheet) or max y (decrementing down the sheet).

    Returns
    -------
    None.
    '''
    assert bin_opt in ['SW', 'HW'], f"bin_opt must be 'SW' or 'HW'. '{bin_opt}' is invalid input"
    colors = [
        '#ffffff', '#ffe119', '#4363d8', '#e6194b', 
        '#f58231', '#911eb4', '#46f0f0', '#f032e6', 
        '#bcf60c', '#fabebe', '#008080', '#e6beff', 
        '#9a6324', '#fffac8', '#aaffc3', 
        '#808000', '#ffd8b1', '#808080', 
    ]

    bin_info = {}
    # bin_info[1] = {"name": "GOOD_BIN1", "count": 0}
    
    for die in die_info.values():
        bin_num = die['sbin_num'] if bin_opt == 'SW' else die['hbin_num']
        bin_name = die['sbin_name'] if bin_opt == 'SW' else die['hbin_name']
        if bin_num not in bin_info:
            bin_info[bin_num] = {"name": bin_name, "count": 0}
        bin_info[bin_num]['count'] += 1
        
    x_range, y_range = get_xy_ranges(die_info)
    x_min = min(x_range)
    y_min = min(y_range)
    print("x_range:", x_range, ", y_range:", y_range)
    print("x_min:", x_min, ", y_min:", y_min)
    
    with xlsxwriter.Workbook(xlsx_fp) as wb:
        ws = wb.add_worksheet("wafer map")
        ws.set_zoom(70)
        ws.freeze_panes(1,1)
    
        cell_format = wb.add_format()
        cell_format.set_center_across()
        for i, x in enumerate(x_range):
            ws.write(0, i+2, 'X{}'.format(x), cell_format)
        
        if top_is_y_min:
            for i, y in enumerate(y_range):
                ws.write(i+2, 0, 'Y{}'.format(y), cell_format)
        else:
            for i, y in enumerate(reversed(y_range)):
                ws.write(i+2, 0, 'Y{}'.format(y), cell_format)
                
        bin_colors = {} # {<soft_bin>: <color_str>}
        if good_bin in bin_info:
            bin_colors[good_bin] = "#3cb44b" # bin 1 always green
        
        for die in die_info.values():
            die['format'] = wb.add_format()
            
        # write_wafermap_border(die_info, ws)
        
        for (x, y), die in die_info.items():
            bin_num = die["sbin_num"] if bin_opt == 'SW' else die["hbin_num"]
            cell_format = die["format"]
            if bin_num not in bin_colors:
                bin_colors[bin_num] = colors.pop() if len(colors) > 1 else colors[0]
            cell_format.set_bg_color(bin_colors[bin_num])
            cell_format.set_center_across()
            if top_is_y_min:
                ws.write(y-y_min+2, x-x_min+2, bin_num, cell_format)
            else:
                ws.write(len(y_range)-(y-y_min-1), x-x_min+2, bin_num, cell_format)
                
            
        header = ["Bin Code", "Name", "%Yield", "Count"]
        for i, string in enumerate(header):
            ws.write(2, len(x_range) + 4 + i, string)
        print("bin codes:", bin_colors.keys())
        if bin_colors.keys():
            for y, bin_num in enumerate(sorted(bin_colors.keys())):
                cell_format = wb.add_format()
                cell_format.set_bg_color(bin_colors[bin_num])
                cell_format.set_center_across()
                ws.write(y + 3, len(x_range) + 4, bin_num, cell_format)
                try:
                    ws.write(y + 3, len(x_range) + 5, bin_info[bin_num]["name"])
                except KeyError:
                    print(f"key error ({bin_num})")
                percent_yield = 100 * bin_info[bin_num]["count"] / len(die_info)
                ws.write(y + 3, len(x_range) + 6, "{:.2f}".format(percent_yield))
                ws.write(y + 3, len(x_range) + 7, bin_info[bin_num]["count"])
                
    # open file in excel
    # add quotes around any directory name with spaces, or system command wont work
    if open_xlsx:
        xlsx_fp = os.path.normpath(xlsx_fp)
        splits = xlsx_fp.split('\\')
        tmp = ""
        for split in splits:
            if  ' ' in split:
                split = ('"' + split + '"')
            tmp += (split + "\\")
        xlsx_fp = tmp[0:-1]
        os.system(xlsx_fp)

if __name__ == "__main__":
    xlsx_fp = r'C:/Users/dkane/OneDrive - Presto Engineering/Documents/python_scripts/wafer-map/dummy_wafer_map.xlsx'
    die_info = {
        (3,3) : {'sbin_num' : 1,        'sbin_name' : 'GOOD_BIN1',      'hbin_num' : 1,     'hbin_name' : 'GOOD_HW_BIN1'},
        (4,3) : {'sbin_num' : 5000,     'sbin_name' : 'DUMMY_FAIL1',    'hbin_num' : 5,     'hbin_name' : 'DUMMY_HW_FAIL1'},
        (5,3) : {'sbin_num' : 1,        'sbin_name' : 'GOOD_BIN1',      'hbin_num' : 1,     'hbin_name' : 'GOOD_HW_BIN1'},
        (3,4) : {'sbin_num' : 20100,    'sbin_name' : 'DUMMY_FAIL2',    'hbin_num' : 20,    'hbin_name' : 'DUMMY_HW_FAIL2'},
        (4,4) : {'sbin_num' : 1,        'sbin_name' : 'GOOD_BIN1',      'hbin_num' : 1,     'hbin_name' : 'GOOD_HW_BIN1'},
        (5,4) : {'sbin_num' : 5000,     'sbin_name' : 'DUMMY_FAIL1',    'hbin_num' : 5,     'hbin_name' : 'DUMMY_HW_FAIL1'},
        (3,5) : {'sbin_num' : 1,        'sbin_name' : 'GOOD_BIN1',      'hbin_num' : 1,     'hbin_name' : 'GOOD_HW_BIN1'},
        (4,5) : {'sbin_num' : 20100,    'sbin_name' : 'DUMMY_FAIL2',    'hbin_num' : 20,    'hbin_name' : 'DUMMY_HW_FAIL2'},
        (5,5) : {'sbin_num' : 20100,    'sbin_name' : 'DUMMY_FAIL2',    'hbin_num' : 20,    'hbin_name' : 'DUMMY_HW_FAIL2'},
    }
    wafer_map(die_info, xlsx_fp, bin_opt='SW', open_xlsx=True)

