# Includes
import clr
from Spotfire.Dxp.Data import DataColumn, TagsColumn, DataManager
from Spotfire.Dxp.Data import DataPropertyClass, DataType, DataValueCursor, IDataColumn, IndexSet
from Spotfire.Dxp.Data import RowSelection 
import datetime

def  tag_Marked_Items(tableName, colName, tagName):
	# Return: 			No return object(s)
	# Last updated:		6/30/18 
	# Created by: 	    Tyler Hunt
	#
	# 1.0 - Create a reference to the targeted column
	# 2.0 - Creates a filter to only the marked row in the table
	# 3.0 - Finally, it tags the marked item(s) in the table
	# 	3.1 - Tag the filtered data set with the 
	# 	3.2 - Print confirmation message(s)

	# 1.0 - Create reference to the tags column 
	tagsCol = tableName.Columns[colName].As[TagsColumn]()

	# 2.0 - Get the marked rows
	filteringRowSelection = Document.ActiveMarkingSelectionReference.GetSelection(tableName)
	filteredSet = IndexSet(filteringRowSelection.AsIndexSet())

	timeUpdated = datetime.datetime.now().strftime("%m-%d-%Y %H:%M:%S")

	# 3.0 - Check if index set is empty
	if filteredSet.Count > 0:
		# 3.1 - Tag the filtered data set with the 
		tagsCol.Tag(tagName,RowSelection(filteredSet))

		# 3.2 - Print confirmation message
		print('Tag(s) has been removed successfully @ ' + timeUpdated)
	else:
		# 3.2 - Print confirmation message
		print('No tag(s) have been created @ ' + timeUpdated)

def get_BhaDeSelectedwells(bhaTable):
	# Return: 			list of bha's selected
	# Last updated:		6/30/18 
	# Created by: 	    Tyler Hunt
	#
	# 1.0 - Create list to hold bha marked pn
	# 2.0 - Create Cursor references to the PN and Tag  in the wells tables
	# 3.0 - Iterate through table column rows to retrieve the values
	# 	3.1 - Get list of bha pn's that are not selected

	# 1.0 - Create list to hold bha marked pn
	bha_deselectedPnList = []

	# 2.0 - Create Cursor references to the PN and Tag  in the wells tables
	cursor_BhaPN = DataValueCursor.CreateFormatted(bhaTable.Columns['propertynumber'])
	cursor_BhaTags = DataValueCursor.CreateFormatted(bhaTable.Columns['Selected'])

	# 3.0 - Iterate through table column rows to retrieve the values
	for row in bhaTable.GetRows(cursor_BhaPN, cursor_BhaTags):
		pn = cursor_BhaPN.CurrentValue
		tag = cursor_BhaTags.CurrentValue

		# 3.1 - Get list of bha pn's that are not selected
		if pn <> str.Empty and tag != 'Selected':
			bha_deselectedPnList.append(str(pn))

	print(bha_deselectedPnList)
	return bha_deselectedPnList

def untag_bhaSelectedItems(columnName, tableName, bha_deselectedWells):
	# Return: 			list of bha's selected
	# Last updated:		6/30/18 
	# Created by: 	    Tyler Hunt
	#
	# 1.0 - Create cursor to the bha selected marker
	# 2.0 - Create references to the bha table row count, rows to include and rows to mark
	# 3.0 - Loop thru the selected wells and add bhaSelected tag to wellsTable
	# 	3.1 - Loop through the selcted well bha list
	# 	3.2 - Loop through the wellsTable and mark the bha selected pn in wellsTable
	# 4.0 - Reference to the Tag Column
	# 5.0 - Get the marked rows
	# 	5.1 - Tag the filtered data set with the 
	# 	5.2 - Print confirmation message(2)

	# 1.0 - Create cursor to the bha selected marker
	cursor_WellsPN = DataValueCursor.CreateFormatted(wellsTable.Columns['alt_well_pn'])
	
	# 2.0 - Create references to the bha table row count, rows to include and rows to mark
	rowCount = wellsTable.RowCount
	rowsToInclude = IndexSet(rowCount,True)
	rowsToMark = IndexSet(rowCount,False)

	# 3.0 - Loop thru the selected wells and add bhaSelected tag to wellsTable
	if not bha_deselectedWells:
		pass
	else:
		# 3.1 - Loop through the selcted well bha list
		for bhaPN in bha_deselectedWells:

			# 3.2 - Loop through the wellsTable and mark the bha selected pn in wellsTable
			for row in wellsTable.GetRows(cursor_WellsPN):
				rowIndex = row.Index 
				wellsPN = cursor_WellsPN.CurrentValue

				# 3.3 - Mark well in wells table 
					# TODO: Might need to check the Selcted status????
				if wellsPN == bhaPN:
					rowsToMark.AddIndex(rowIndex)
					Document.ActiveMarkingSelectionReference.SetSelection(RowSelection(rowsToMark), wellsTable)

		# 4.0 - Reference to the Tag Column
		tagsCol = tableName.Columns[columnName].As[TagsColumn]()

		# Tag nname
		tagName =''

		# Time updated
		timeUpdated = datetime.datetime.now().strftime("%m-%d-%Y %H:%M:%S")

		# 5.0 - Get the marked rows
		filteringRowSelection = Document.ActiveMarkingSelectionReference.GetSelection(tableName)
		filteredSet = IndexSet(filteringRowSelection.AsIndexSet())

		# 5.0 - Check if index set is empty
		if filteredSet.Count != 0:
			# 5.1 - Tag the filtered data set with the 
			tagsCol.Tag(tagName,RowSelection(filteredSet))

			# 5.2 - Print confirmation message
			print('Bha Tag(s) have been removed successfully @ ' + timeUpdated)
		else:
			# 5.2 - Print confirmation message
			print('No Bha tag(s) have been removed @ ' + timeUpdated)
	


def clear_Markings(markingName, tableName):
	# Return: 			No return object(s)
	# Last updated:		6/30/18 
	# Created by: 	    Tyler Hunt
	#
	# 1.0 - Create reference to the markings
	# 2.0 - Create index set on the current table reference
	# 3.0 - Creates a cursor for the marked items in the table

	# 1.0 - Create reference to the markings
	marking = Application.GetService[DataManager]().Markings[markingName]

	# 2.0 - Create index set on the current table reference
	selectRows = IndexSet(tableName.RowCount, False)

	# 3.0 - Clar the markingon the table
	marking.SetSelection(RowSelection(selectRows),tableName)
	print('Markings have been cleared')

# Main - Remove Wells Selection Tags ---------------------------------------------------

# Last updated:		6/30/18 
# Created by: 	    Tyler Hunt
#
# 1.0 - Tag marked items
# 2.0 - Clear marked items

# Table references variables 
wellsTable = Document.Data.Tables['similarWells']
bhaTable = Document.Data.Tables['bhaInfo']

# Column, Marker and Tag reference variables
colName = 'Selected'
markerName = 'Deselect Bha'
tagName =  ''

# 1.0 - Tag marked items
tag_Marked_Items(bhaTable, colName, tagName)

# 2.0 - Clear marked items
clear_Markings(markerName, bhaTable)

# 3.0 - Deselct
bha_deselectWells = get_BhaDeSelectedwells(bhaTable)

# 4.0 - UnTag
untag_bhaSelectedItems('bhaSelected', wellsTable, bha_deselectWells)

# 5.0 - Clear marked items
clear_Markings(markerName, bhaTable)
