#!/usr/bin/env python
from openpyxl import load_workbook
from argparse import ArgumentParser
import glob
import os

parser = ArgumentParser(description="Add L3 to 9.32 Report")
parser.add_argument("files", nargs="*", help="Supply both L3 file and 9.32 file as Excel document on the command line")


ns = parser.parse_args()

if ns.files:
    if "l3" in ns.files[0].lower() and "9.32" in ns.files[1].lower():
        lswb = load_workbook(ns.files[0])
        nwb = load_workbook(ns.files[1])
    elif "l3" in ns.files[1] and "9.32" in ns.files[0].lower():
        l3wb = load_workbook(ns.files[1])
        nwb = load_workbook(ns.files[0])
else:
    #We didn't pass any files on the command line. Figure it out. 
    os.chdir(os.path.dirname(__file__))
    files = [fn for fn in glob.glob("*.xlsx") if not fn.startswith("~")]
    l3files = [fn for fn in files if "l3" in fn.lower()]
    nfiles = [fn for fn in files if "9.32" in fn.lower()]
    exit = False
    if len(l3files) != 1:
        print("Too many or too few l3 files")
        exit = True
    if len(nfiles) != 1:
        print("Too many or too few 9.32 files")
        exit = True
    if exit:
        sys.exit(1)
    l3wb = load_workbook(l3files[0])
    nwb = load_workbook(nfiles[0])

print l3wb
print nwb

