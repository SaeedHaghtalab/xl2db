#!/bin/env python
#
#  Copyright (c) 2019  European Spallation Source ERIC
#
#  The program is free software: you can redistribute
#  it and/or modify it under the terms of the GNU General Public License
#  as published by the Free Software Foundation, either version 2 of the
#  License, or any newer version.
#
#  This program is distributed in the hope that it will be useful, but WITHOUT
#  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
#  FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for
#  more details.
#
#  You should have received a copy of the GNU General Public License along with
#  this program. If not, see https://www.gnu.org/licenses/gpl-2.0.txt
#
#
#   author  : Saeed Haghtalab
#   email   : saeed.haghtalab@ess.eu
#   date    : May 28 16:15:42 CEST 2020
#   version : 0.0

import xlrd

headerRow = 0
recNameCol = 0
recTypeCol = 1
outRecTypes = {'aao', 'ao', 'bo', 'mbbo', 'mbboDirect', 'calcout', 'longout', 'stringout'}
inRecTypes = {'aai', 'ai', 'bi', 'mbbi', 'mbbiDirect', 'longin', 'stringin'}

recNamePref = '$(P)$(R)'
fieldInpOutPref = '@asyn($(PORT),$(ADDR),$(TIMEOUT))'

rec = []


def main():
    xlDbLoc = "../template/cromeDb.xlsx"
    workbook = xlrd.open_workbook(xlDbLoc)

    for db in workbook.sheets():
        if db.name == 'Lists':
            continue
        db = workbook.sheet_by_index(0)
        print("Generating epics database for:",db.name)
        print(db.nrows - 1, "Records detected")
        dblist = []
        for recrow in range(headerRow + 1, db.nrows):
            recName = recNamePref + db.cell(recrow, recNameCol).value
            recType = db.cell(recrow, recTypeCol).value
            if not recName or not recType:
                print('Blank record name or record type in row ' + str(recrow+1) + '! Skipping ...')
                continue
            rec = 'record(' + recType + ', "' + recName + '"){'
            dblist.append(rec)
            for reccol in range(recTypeCol + 1, db.ncols):
                fieldName = db.cell(headerRow, reccol).value
                fieldVal = str(db.cell(recrow, reccol).value)
                if fieldVal and fieldName:
                    if fieldName == 'INP/OUT':
                        fieldVal = fieldInpOutPref + fieldVal
                        if recType in outRecTypes:
                            fieldName = 'OUT'
                        elif recType in inRecTypes:
                            fieldName = 'INP'
                        else:
                            print('ERROR: No record type found correspond to "INP/OUT" field')
                            exit()
                    field = '    field(' + fieldName + ', "' + fieldVal + '")'
                    dblist.append(field)
            dblist.append('}')
        f = open("../" + db.name,"w+")
        for line in range (len(dblist)):
            f.write(dblist[line] + '\n')

if __name__=="__main__":
    main()
