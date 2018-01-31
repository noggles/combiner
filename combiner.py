# combiner.py
'''
Written in Python 2.7
This script was designed to combine all of the quantification tables output by the Phenom SEM.
The order of the oxides has been set to match the format of typical mineral identication tables.

For easiest use, place this script in the same folder as your SEM Report folder. When running
this script, simply enter the name of the report folder at the first question.

You may or may not be able to overwrite previous excel documents with the same name as your output.
Executing this script as a superuser/admin may allow you to overcome this. If not, just delete any
other 'SEM_data.xls' file or rename the output file at the bottom of this script.

Please ensure you are using the correct version of Python. This script does not work in python 3.
'''

import os
import csv
import xlwt
import re

# FUNCTIONS
def list_files(filepath):
    '''Scans the report directory and collects a list of the names of all files within'''
    r = []
    subdirs = [x[0] for x in os.walk(filepath)]
    for subdir in subdirs:
        files = os.walk(subdir).next()[2]
        if (len(files) > 0):
            for file in files:
                if file[-1] == "v":
                    r.append(subdir + "/" + file)
    return r


def harvester(filepath, nitro=True):
    '''
    Reads through a data file and harvests all the information. Outputs in a form
    useful for the script procedure.

    :param filepath:    location of file to harvest data from
    :param type:        <string>
    :param nitro:       set to True if you would like NiO to be included in the
                        MID compatible output sheet. default is True
    :param type:         <boolean>
    :return:            data for script procedure
    :rtype:             <array>
    '''
    filename = filepath[:-20]
    list1 = filepath.split('/')         # grab the filename from filepath
    filename = list1[0][-14:]
    filename = list1[-2]
    if 'spot' in filename:              # check if image_spot
        atn = []
        els = []
        eln = []
        atc = []
        wtc = []
        oxn = []
        swp = []
        with open(filepath) as csvfile:             # open the image_spot file
            for row in csv.reader(csvfile):
                airlock = []
                atn.append(row[0])
                els.append(row[1])
                eln.append(row[2])
                atc.append(row[3])
                wtc.append(row[4])
                try:                                # replace element descriptors with oxide names
                    airlock.append(row[5])
                    if airlock[0] == 'Si':
                        oxn.append('SiO')
                    elif airlock[0] == 'Al':
                        oxn.append('Al2O3')
                    elif airlock[0] == 'Cr':
                        oxn.append('Cr2O3')
                    elif airlock[0] == 'Ti':
                        oxn.append('TiO2')
                    elif airlock[0] == 'K':
                        oxn.append('K2O')
                    elif airlock[0] == 'Ca':
                        oxn.append('CaO')
                    elif airlock[0] == 'Na':
                        oxn.append('Na2O')
                    elif airlock[0] == 'Ni':
                        if nitro == True:
                            oxn.append('NiO')
                    elif airlock[0] == 'Fe':
                        oxn.append('FeO')
                    elif airlock[0] == 'Mg':
                        oxn.append('MgO')
                    elif airlock[0] == 'Mn':
                        oxn.append('MnO')
                    elif airlock[0] == 'P':
                        oxn.append('P2O5')
                    else:
                        oxn.append(str(row[5])+' oxide')
                    swp.append(row[6])
                except IndexError:
                    pass
        atn = atn[1:]       # shave labels off each column
        els = els[1:]
        eln = eln[1:]
        atc = atc[1:]
        wtc = wtc[1:]
        oxn = oxn[1:]
        swp = swp[1:]
        temp2=[]
        temp3=[]
        try:                        # rearrange oxides into correct order
            A = oxn.index('SiO')
            temp2.append(oxn[A])
            temp3.append(swp[A])
            del oxn[A]
            del swp[A]
        except:
            pass
        try:
            B = oxn.index('TiO2')
            temp2.append(oxn[B])
            temp3.append(swp[B])
            del oxn[B]
            del swp[B]
        except:
            pass
        try:
            C = oxn.index('Cr2O3')
            temp2.append(oxn[C])
            temp3.append(swp[C])
            del oxn[C]
            del swp[C]
        except:
            pass
        try:
            D = oxn.index('Al2O3')
            temp2.append(oxn[D])
            temp3.append(swp[D])
            del oxn[D]
            del swp[D]
        except:
            pass
        try:
            E = oxn.index('FeO')
            temp2.append(oxn[E])
            temp3.append(swp[E])
            del oxn[E]
            del swp[E]
        except:
            pass
        try:
            F = oxn.index('MnO')
            temp2.append(oxn[F])
            temp3.append(swp[F])
            del oxn[F]
            del swp[F]
        except:
            pass
        try:
            G = oxn.index('NiO')
            temp2.append(oxn[G])
            temp3.append(swp[G])
            del oxn[G]
            del swp[G]
        except:
            pass
        try:
            H = oxn.index('MgO')
            temp2.append(oxn[H])
            temp3.append(swp[H])
            del oxn[H]
            del swp[H]
        except:
            pass
        try:
            I = oxn.index('CaO')
            temp2.append(oxn[I])
            temp3.append(swp[I])
            del oxn[I]
            del swp[I]
        except:
            pass
        try:
            J = oxn.index('Na2O')
            temp2.append(oxn[J])
            temp3.append(swp[J])
            del oxn[J]
            del swp[J]
        except:
            pass
        try:
            K = oxn.index('K2O')
            temp2.append(oxn[K])
            temp3.append(swp[K])
            del oxn[K]
            del swp[K]
        except:
            pass
        try:
            K = oxn.index('P2O5')
            temp2.append(oxn[K])
            temp3.append(swp[K])
            del oxn[K]
            del swp[K]
        except:
            pass
        # prepare final output
        temp2 = temp2 + oxn
        temp3 = temp3 + swp
        datapiece = [atn,els,eln,atc,wtc,temp2,temp3,filename]
        return datapiece
    else:
        return 'failure'

###############################################################################################
# Script procedure
print 'IT IS COMBINE TIME!!!'
#print ("Windows filepath guide: C:\Users\...\Export")
#print ("Linux filepath guide: /folder/subfolder/.../Export")

filepath = raw_input("enter filepath to SEM report folder here:")
filenames = list_files(filepath)
pile = []                               # largest data structure in script, contains all
spot_count = 0                          # how many image_spots are detected
fail_count = 0                          # how many other data styles are detected
for file in filenames:                  # iterate through each file
    datapiece = harvester(file)         # harvest data from file
    if datapiece == "failure":          # check if image_spot or not
        fail_count = fail_count + 1
        continue
    else:
        pile.append(datapiece)          # add the image_spot data to the pile
        spot_count = spot_count + 1
print 'image_spots detected: '+str(spot_count)
print 'maps/line_scans skipped: '+str(fail_count)
i=0
for datapiece in pile:
    if i==0:
        prev = datapiece
    else:
        prev = pile[i-1]
        datapiece[7]

# Instructions to generate general SEM data spreadsheet
book = xlwt.Workbook(encoding="utf-8")      # open new excel workbook
sheet1 = book.add_sheet("SEM data")         # create first sheet
sheet1.write(0,0, 'atomic num')
sheet1.write(0,1, 'atomic sym')
sheet1.write(0,2, 'element name')
sheet1.write(0,3, 'atomic conc.')
sheet1.write(0,4, 'weight conc.')
sheet1.write(0,5, 'oxide name')
sheet1.write(0,6, 'stoich. conc.')
sheet1.write(0,7, 'image.spot')

samples = {}
for datapiece in pile:
    key = datapiece[7]
    key_pieces = re.findall(r'\d+', str(key))
    if len(key_pieces[1])==1:
        key_pieces[1] = '0'+ str(key_pieces[1])
    key_number = str(key_pieces[0])+'.'+str(key_pieces[1])
    key_number = float(key_number)
    samples[key_number] = datapiece[:7]
i = 1
for key in sorted(samples):
    datapiece = samples[key]
    L = len(datapiece[0])
    j = 0
    while j < L:
        sheet1.write(i, 0, datapiece[0][j])
        sheet1.write(i, 1, datapiece[1][j])
        sheet1.write(i, 2, datapiece[2][j])
        sheet1.write(i, 3, datapiece[3][j])
        sheet1.write(i, 4, datapiece[4][j])
        try:
            sheet1.write(i, 5, datapiece[5][j])
            sheet1.write(i, 6, datapiece[6][j])
        except IndexError:
            pass
        sheet1.write(i, 7, key)
        j = j + 1
        i = i + 1


# Instructions to generate a mineral identification compatible spreadsheet
samples={}
y=0
x=1
for datapiece in pile:
    datapiece = [datapiece[5], datapiece[6], datapiece[7]]
    temp3 = []
    temp4 = []
    if 'SiO' in datapiece[0]:
        A = datapiece[0].index('SiO')
        temp3.append(datapiece[0][A])
        temp4.append(datapiece[1][A])
        del datapiece[0][A]
        del datapiece[1][A]
    else:
        temp3.append('SiO')
        temp4.append(0.)
    if 'TiO2' in datapiece[0]:
        A = datapiece[0].index('TiO2')
        temp3.append(datapiece[0][A])
        temp4.append(datapiece[1][A])
        del datapiece[0][A]
        del datapiece[1][A]
    else:
        temp3.append('TiO2')
        temp4.append(0.)
    if 'Cr2O3' in datapiece[0]:
        A = datapiece[0].index('Cr2O3')
        temp3.append(datapiece[0][A])
        temp4.append(datapiece[1][A])
        del datapiece[0][A]
        del datapiece[1][A]
    else:
        temp3.append('Cr2O3')
        temp4.append(0.)
    if 'Al2O3' in datapiece[0]:
        A = datapiece[0].index('Al2O3')
        temp3.append(datapiece[0][A])
        temp4.append(datapiece[1][A])
        del datapiece[0][A]
        del datapiece[1][A]
    else:
        temp3.append('Al2O3')
        temp4.append(0.)
    if 'FeO' in datapiece[0]:
        A = datapiece[0].index('FeO')
        temp3.append(datapiece[0][A])
        temp4.append(datapiece[1][A])
        del datapiece[0][A]
        del datapiece[1][A]
    else:
        temp3.append('FeO')
        temp4.append(0.)
    if 'MnO' in datapiece[0]:
        A = datapiece[0].index('MnO')
        temp3.append(datapiece[0][A])
        temp4.append(datapiece[1][A])
        del datapiece[0][A]
        del datapiece[1][A]
    else:
        temp3.append('MnO')
        temp4.append(0.)
    if 'NiO' in datapiece[0]:
        A = datapiece[0].index('NiO')
        temp3.append(datapiece[0][A])
        temp4.append(datapiece[1][A])
        del datapiece[0][A]
        del datapiece[1][A]
    else:
        temp3.append('NiO')
        temp4.append(0.)
    if 'MgO' in datapiece[0]:
        A = datapiece[0].index('MgO')
        temp3.append(datapiece[0][A])
        temp4.append(datapiece[1][A])
        del datapiece[0][A]
        del datapiece[1][A]
    else:
        temp3.append('MgO')
        temp4.append(0.)
    if 'CaO' in datapiece[0]:
        A = datapiece[0].index('CaO')
        temp3.append(datapiece[0][A])
        temp4.append(datapiece[1][A])
        del datapiece[0][A]
        del datapiece[1][A]
    else:
        temp3.append('CaO')
        temp4.append(0.)
    if 'Na2O' in datapiece[0]:
        A = datapiece[0].index('Na2O')
        temp3.append(datapiece[0][A])
        temp4.append(datapiece[1][A])
        del datapiece[0][A]
        del datapiece[1][A]
    else:
        temp3.append('Na2O')
        temp4.append(0.)
    if 'K2O' in datapiece[0]:
        A = datapiece[0].index('K2O')
        temp3.append(datapiece[0][A])
        temp4.append(datapiece[1][A])
        del datapiece[0][A]
        del datapiece[1][A]
    else:
        temp3.append('K2O')
        temp4.append(0.)
    if 'P2O5' in datapiece[0]:
        A = datapiece[0].index('P2O5')
        temp3.append(datapiece[0][A])
        temp4.append(datapiece[1][A])
        del datapiece[0][A]
        del datapiece[1][A]
    else:
        temp3.append('P2O5')
        temp4.append(0.)
    temp3.append('other')
    others = [float(i) for i in datapiece[1]]
    sum_others = sum(others)
    temp4.append(sum_others)
    key = datapiece[2]
    key_pieces = re.findall(r'\d+', str(key))
    if len(key_pieces[1])==1:
        key_pieces[1] = '0'+ str(key_pieces[1])
    key_number = str(key_pieces[0])+'.'+str(key_pieces[1])
    key_number = float(key_number)
    samples[key_number]=temp4
    y=y+1
# sort samples
sample_count = len(samples)

if sample_count>250:
    sheets = sample_count/250
    if sample_count%250>0:
        sheets = sheets +1

sheet = book.add_sheet('MID 1')         # add first MID sheet
sheet.write(0,0, 'image.spot')
h = 1
for value in temp3:
    sheet.write(h, 0, temp3[h-1])
    h = h + 1
x=0
p=0
for key in sorted(samples):
    i = 1
    if p == 250:                            # if sample count exceeds excels 256 col limit...
        sheet_name = 'MID ' + str(i + 1)    # ...create a new sheet
        sheet = book.add_sheet(sheet_name)
        sheet.write(0, 0, 'image.spot')
        h=1
        for value in temp3:
            sheet.write(h, 0, temp3[h-1])
            h=h+1
        p=0
        i=i+1
    sheet.write(0, p+1, key)  # Write in sample name
    L = len(temp3)
    y = 1
    j = 0
    while j < L:  # Write in data column for each image spot
        dataset = samples[key]
        sheet.write(y, p+1, dataset[j])
        j = j + 1
        y = y + 1
    x = x + 1
    p = p + 1

book.save("SEM_data.xls")
print 'number of MID sheets: ' +str(sheets) + '  (max 250 image_spots/sheet)'
print "SUCCESS! SEM_data.xls has been generated."
