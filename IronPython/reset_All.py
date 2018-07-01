#Includes
import clr
from Spotfire.Dxp.Data import DataColumn, TagsColumn, DataManager
from Spotfire.Dxp.Data import DataPropertyClass, DataType, DataValueCursor, IDataColumn, IndexSet
from Spotfire.Dxp.Data import RowSelection 
import datetime

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
	print('The marking ' + str(markingName) + ' have been cleared')


def remove_allTags(tableName, colName):
	# Last updated:		6/30/18 
	# Created by: 	    Tyler Hunt
	# Return: 			none
	#
	# 1.0 - Create cursor to the reference column
	# 2.0 - Create references to the row count and rows to mark
	# 3.0 - Loop through the table and mark the selected pn in table

	# 1.0 - Create cursor to the reference column
	cursor_1 = DataValueCursor.CreateFormatted(tableName.Columns[colName])
	
	# 2.0 - Create references to the row count and rows to mark
	rowCount = tableName.RowCount
	rowsToMark = IndexSet(rowCount,False)

	# 3.0 - Loop through the table and mark the selected pn in table
	for row in tableName.GetRows(cursor_1):
		rowIndex = row.Index 
		select = cursor_1.CurrentValue

		# 3.1 - Mark well(s)
		if select == 'Selected':
			rowsToMark.AddIndex(rowIndex)
			Document.ActiveMarkingSelectionReference.SetSelection(RowSelection(rowsToMark), tableName)

	# 4.0 - Reference to the Tag Column
	tagsCol = tableName.Columns[colName].As[TagsColumn]()

	# Tag nname - Clear marking
	tagName =''

	# Time updated
	timeUpdated = datetime.datetime.now().strftime("%m-%d-%Y %H:%M:%S")

	# 5.0 - Get the marked rows
	filteringRowSelection = Document.ActiveMarkingSelectionReference.GetSelection(tableName)
	filteredSet = IndexSet(filteringRowSelection.AsIndexSet())

	# 5.0 - Check if index set is empty
	if filteredSet.Count > 0:
		# 5.1 - Tag the filtered data set with the 
		tagsCol.Tag(tagName,RowSelection(filteredSet))

		# 5.2 - Print confirmation message
		print('All tag(s) have been successfully removed @ ' + timeUpdated)
	else:
		# 5.2 - Print confirmation message
		print('No tag(s) have been removed @ ' + timeUpdated)

# Main ------------------------------------------------------
# Last updated:		7/1/18 
# Created by: 	    Tyler Hunt
#
# 1.0 - Get all bha well in in bhaInfo table
# 2.0 - Remove all bhaInfo table markings

# Column and marking name variable
colName_1 = 'Selected'
colName_2 = 'bhaSelected'

markerName_1 = 'Select Well'
markerName_2 = 'Deselect Well'
markerName_3 = 'Select Bha'
markerName_4 = 'Deselect Bha'

# Table references variables 
wellsTable = Document.Data.Tables['similarWells']
bhaTable = Document.Data.Tables['bhaInfo']

# 1.0 - Tag and deselect all bhaTable wells
remove_allTags(wellsTable, colName_1)
remove_allTags(wellsTable, colName_2)
remove_allTags(bhaTable, colName_1)

# 2.0 - Remove all bhaInfo table markings
clear_Markings(markerName_1, wellsTable)
clear_Markings(markerName_2, wellsTable)

clear_Markings(markerName_3, bhaTable)
clear_Markings(markerName_4, bhaTable)
