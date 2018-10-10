#!/usr/bin/env python
from openpyxl import load_workbook
from argparse import ArgumentParser
import glob
import os

parser = ArgumentParser(description="Add L3 to 9.32 Report")
parser.add_argument("files", nargs="*", help="Supply both L3 file and 9.32 file as Excel document on the command line")


ns = parser.parse_args()

if ns.files:
    if "ls" in ns.files[0].lower() and "9.32" in ns.files[1].lower():
        lswb = load_workbook(ns.files[0])
        nwb = load_workbook(ns.files[1])
    elif "ls" in ns.files[1] and "9.32" in ns.files[0].lower():
        lswb = load_workbook(ns.files[1])
        lswb = load_workbook(ns.files[0])
else:
    #We didn't pass any files on the command line. Figure it out. 
    os.chdir(os.path.dirname(__file__))
    files = glob.glob("*.xlsx")
    print files