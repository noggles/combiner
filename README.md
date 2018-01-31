# combiner

Designed for compiling all output data tables from a Phenom SEM report into a single Excel workbook, in a format compatible with mineral identification sheets.

Written in Python 2.7
Required python modules: xlwt, re, os, csv

For easiest use, place combiner.py in the same directory as the SEM report. Enter the name of the report at the prompt. Otherwise provide the filepath to the report.

For reports containing more than 256 image spots, multiple MID sheets will be generated consisting of 250 image spots each. This is due to Excels formatting limit of 256 columns.
