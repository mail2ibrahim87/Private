"""Standalone chemical cleaning / decontamination volume template (row layout, units stacked).

Usage:  python3 tools/build_volume_template.py Cleaning_Decon_Volume_Template.xlsx
"""
import sys

from openpyxl import Workbook

import volume_rows as vr


def build(out):
    wb = Workbook()
    wb.remove(wb.active)
    vr.build_lookup(wb)
    vr.build_settings(wb, "", "Settings")
    vol = vr.build_volume(wb, "", "Volume Calculation", "Settings")
    vr.build_summary(wb, "", "Summary", vol)
    vr.build_chemicals(wb, "", "Chemicals")
    wb._sheets = [wb[n] for n in ("Settings", "Volume Calculation", "Summary", "Chemicals", "Lookup")]
    wb.active = 1
    wb.save(out)
    print("saved", out)


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "Cleaning_Decon_Volume_Template.xlsx")
