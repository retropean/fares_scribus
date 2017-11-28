#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ABOUT THIS SCRIPT (modified by J.-P. Gendner June, 15 2014):

Import CSV data files as tables into Scribus

1. You may create a style to be applied to all cells of the table

2. Create a frame at the position (upper left corner) 
and of the size (height and width) you want for your table

3. Make sure it is selected

4. Execute this script:

  You will first be prompted for the csv data filename.
  ->  The number of columns and row will be displayed.
  You will then be prompted for the width of each column in mm.
  ->  The default value given is calculated to respect the width of the
table.
  Then you will be prompted for the height of each row.
  ->  The default value is calculated to respect the height of the table.
  Then you may enter the style name you want to be applied to each cell.
  Finally, you may enter the Top distance for all text.

5. The data from the csv file will be imported and a table of
textboxes will be drawn on the page.

LIMITATIONS:

   In Scribus versions 1.3.5svn, when the script ends, you cannot
adjust text, colors, and line features for a group, whereas in 1.3.3.x,
all of these can be done without ungrouping.

HINTS:

Postgresql:
You can easily create a CSV file with a Postgresql database. From
Postgresql,
toggle unaligned output with the '\a' switch, then activate a comma as
a separator with '\f ,' (without apostrophes). Send output to a file
with '\o myfile.csv', then query your database.

Sqlite3:
You can use "sqlite3 -csv" in the command line or ".mode csv" in sqlite's
interactive shell.

############################

LICENSE:

This program is free software; you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation; either version 2 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301,
USA.

Author: Sebastian Stetter

Modifications: Gregory Pittman

please report bugs to: scribusscript-xpvPi5bcW5WP0LLXB2dsMhvVK+yQ3ZXh@public.gmane.org
"""

from __future__ import division
import sys

try:
    # Please do not use 'from scribus import *' . 
    # If you must use a 'from import',
    # Do so _after_ the 'import scribus' 
    # and only import the names you need,
    # such as commonly used constants.
    import scribus
except ImportError,err:
    print "This Python script is written for the Scribus scripting interface."
    print "It can only be run from within Scribus."
    sys.exit(1)

#########################
# YOUR IMPORTS GO HERE  #
#########################
import csv

# get information about the area where the bale should be drawed
def getPosition():
    if scribus.selectionCount() == 1:
        areaname = scribus.getSelectedObject()
        position= scribus.getPosition(areaname)
        vpos = position[1]
        hpos = position[0]
        size = scribus.getSize(areaname)
        width = size[0]
        height = size[1]
        scribus.deleteObject(areaname)
        return vpos, hpos, width, height
        
    else: 
        scribus.messageBox("csv2table", "please select ONE Object to mark the drawing area for the table")
        sys.exit()

# get the csv data
def getCSVdata():
    """opens a csv file, reads it in and returns a 2 dimensional list with the data"""
    csvfile = scribus.fileDialog("csv2table :: open file", "*.csv")
    if csvfile != "":
        try:
            reader = csv.reader(file(csvfile))
            datalist=[]
            for row in reader:
                rowlist=[]
                for col in row:
                    rowlist.append(col)
                datalist.append(rowlist)
            return datalist
        except Exception,  e:
            scribus.messageBox("csv2table", "Could not open file %s"%e)
    else:
        sys.exit

def getDataInformation(list):
    """takes a 2 dimensional list object and returns the numbers of rows and cols"""
    datainfo = dict()
    datainfo["rowcount"]=len(list)
    datainfo["colcount"]= len(list[0])
    return datainfo

def cellsize(areainfo, datainfo):
    """"takes the area and data info and calculates the prper cell size"""
    csize=dict()
    csize["v"]= areainfo["vsize"] / datainfo["rowcount"]
    csize["h"]= areainfo["hsize"] / datainfo["colcount"]
    return csize
    
def main(argv):
	userdim=scribus.getUnit() # get unit and change it to mm
	scribus.setUnit(scribus.UNIT_MILLIMETERS)
	cellwidthleft = 0
	cellwidthright = 0
	pos = getPosition()
	hposition=pos[1]
	vposition=pos[0]
	width = pos[2]
	height = pos[3]
	data = getCSVdata()
	di= getDataInformation(data)
	ncol = len(data[0])
	nrow = len(data)
	scribus.messageBox("Table", "   " + str(ncol) + " columns,    " + str(nrow) + " rows   ")	#jpg
	ColWidthList = []
	TableWidth = 0
	RowHeightList = []
	TableHeight=0
	i=0
	for row in data:
		if i == 0:
			c = 0
			for cell in row:
				ColWidthT = (width-TableWidth)/(ncol-c)
				ColWidth = float(ColWidthT)
				ColWidthList.append(ColWidth)
		TableWidth = TableWidth + ColWidth
		c = c+1
		RowHeightT = (height-TableHeight)/(nrow-i)
		RowHeight = float(RowHeightT)
		RowHeightList.append(RowHeight)
		TableHeight = TableHeight + RowHeight
		i = i+1
	CellsStyle = scribus.valueDialog('Cells Style','Style name or blank?',"")
	TextDist = float(scribus.valueDialog('Text distance','Top distance for text (mm)?',"0"))
	objectlist=[] # here we keep a record of all the created textboxes so we can group them later
	i = 0
	scribus.progressTotal(len(data))
	scribus.setRedraw(False)
	rowindex = 0
	while rowindex < len(data):
		c = 0
		origin_cd = data[rowindex][0].strip()
		origin = data[rowindex][1].strip()
		destination_cd = data[rowindex][2].strip()
		destination = data[rowindex][3].strip()
		fareplan = data[rowindex][4].strip()
		direction = data[rowindex][5].strip()
		fare = data[rowindex][6].strip()
		fare_onboard = data[rowindex][7].strip()
		
		cellsize = ColWidthList[c]
		cellHeight = RowHeightList[i]
		textbox = scribus.createText(hposition, vposition, cellsize, cellHeight) # create a textbox.  
		objectlist.append(textbox)
		if len(CellsStyle) > 0:
			scribus.setStyle(CellsStyle, textbox)
		if TextDist > 0:
			scribus.setTextDistances(0, 0, TextDist, 0, textbox)
		scribus.insertText(origin, 0, textbox) # insert the origin into the textbox.
		vposition = vposition + cellHeight # shift position of cell.  
		c = c + 1
		
		textbox = scribus.createText(hposition, vposition, cellsize, cellHeight) # create a textbox.
		objectlist.append(textbox)
		if len(CellsStyle) > 0:
			scribus.setStyle(CellsStyle, textbox)
		if TextDist > 0:
			scribus.setTextDistances(0, 0, TextDist, 0, textbox) 
		scribus.insertText(destination, 0, textbox) # insert the destination into the textbox.
		vposition=vposition + cellHeight # shift position of cell.  
		c = c + 1
		rowindex = rowindex + 1
		'''
		textbox = scribus.createText(hposition, vposition, cellsize, cellHeight) # create a textbox.
		objectlist.append(textbox)
		if len(CellsStyle) > 0:
			scribus.setStyle(CellsStyle, textbox)
			if TextDist > 0:
				scribus.setTextDistances(0, 0, TextDist, 0, textbox) 
			scribus.insertText(fareplan, 0, textbox) # insert the fareplan into the textbox.
			hposition=hposition+cellsize #move the position for the next cell
			c = c+1
			
		textbox = scribus.createText(hposition, vposition, cellsize, cellHeight) # create a textbox.
		objectlist.append(textbox)
		if len(CellsStyle) > 0:
			scribus.setStyle(CellsStyle, textbox)
			if TextDist > 0:
				scribus.setTextDistances(0, 0, TextDist, 0, textbox) 
			scribus.insertText(fare, 0, textbox) # insert the fare into the textbox.  
			hposition=hposition+cellsize #move the position for the next cell
			c = c+1
		i = i+1
		''''''
        for cell in row:
            cell = cell.strip()
            cellsize=ColWidthList[c]
            cellHeight=RowHeightList[i]
            textbox=scribus.createText(hposition, vposition, cellsize, cellHeight) #create a textbox
            objectlist.append(textbox)
	    if len(CellsStyle) > 0:
              scribus.setStyle(CellsStyle,textbox)
            if TextDist > 0:
              scribus.setTextDistances(0,0,TextDist,0,textbox) 
            scribus.insertText(row[2],0, textbox)#insert the text into the textbox
            hposition=hposition+cellsize #move the position for the next cell
            c=c+1
		vposition = vposition + cellHeight	#set vertical position for next row
		hposition = pos[1] #reset vertical position for next row
		i = i+1
		scribus.progressSet(i)
		'''
	scribus.deselectAll()
	scribus.groupObjects(objectlist)
	scribus.progressReset()
	scribus.setUnit(userdim) # reset unit to previous value
	scribus.docChanged(True)
	scribus.statusMessage("Done")
	scribus.setRedraw(True)

def main_wrapper(argv):
    """The main_wrapper() function disables redrawing, sets a sensible generic
    status bar message, and optionally sets up the progress bar. It then runs
    the main() function. Once everything finishes it cleans up after the main()
    function, making sure everything is sane before the script terminates."""
    try:
        scribus.statusMessage("Importing .csv table...")
        scribus.progressReset()
        main(argv)
    finally:
        # Exit neatly even if the script terminated with an exception,
        # so we leave the progress bar and status bar blank and make sure
        # drawing is enabled.
        if scribus.haveDoc() > 0:
            scribus.setRedraw(True)
        scribus.statusMessage("")
        scribus.progressReset()

# This code detects if the script is being run as a script, or imported as a module.
# It only runs main() if being run as a script. This permits you to import your script
# and control it manually for debugging.
if __name__ == '__main__':
    main_wrapper(sys.argv)