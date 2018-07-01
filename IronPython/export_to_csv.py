import System
import clr
clr.AddReference("System.Windows.Forms")
import sys
import datetime
from sys import exit
from System.Windows.Forms import FolderBrowserDialog, MessageBox, MessageBoxButtons, DialogResult
from Spotfire.Dxp.Data.Export import DataWriterTypeIdentifiers
from System.IO import File, FileStream, FileMode
from Spotfire.Dxp.Application import DocumentMetadata

# Build file name, file type, default path
# 	TODO: get District, Well Name and Rig to build directory path and file name
#	TODO: get sud directory name (district sub-folder)
fileName = 'RoadMap_123'
fileType = '.xls'

fileDetails = fileName + fileType
defaultPath = Document.Properties['exportFileLocation']
savePath = defaultPath + "\\" + fileDetails
print
exportMessage = ''

dialogResult = MessageBox.Show("The RoadMap files are ready to be exported\n\nDo you want to change the save location?\n\nPath: " + savePath 
                , "OSC RoadMap Export Prompt", MessageBoxButtons.YesNoCancel)

if(dialogResult == DialogResult.Yes):
    SaveFile = FolderBrowserDialog()
    SaveFile.ShowDialog()
    savePath = SaveFile.SelectedPath + "\\" + fileDetails

elif(dialogResult == DialogResult.No):
	roadmapTable =  Document.Data.Tables['roadmap']
	writer = Document.Data.CreateDataWriter(DataWriterTypeIdentifiers.ExcelXlsDataWriter)

	stream = File.OpenWrite(savePath)

	allColumnNames = []
	allRows = Document.Data.AllRows.GetSelection(roadmapTable).AsIndexSet()

	for column in roadmapTable.Columns:
		allColumnNames.Add(column.Name)

	writer.Write(stream, roadmapTable, allRows, allColumnNames)
	timeUpdated = datetime.datetime.now().strftime("%m-%d-%Y %H:%M:%S")

	stream.Close()
	stream.Dispose()
	exportMessage = 'The Roadmap files have been successfully exported!\n' + 'Time Saved: ' + timeUpdated + '\nLocation: ' + savePath

else:
	timeUpdated = datetime.datetime.now().strftime("%m-%d-%Y %H:%M:%S")
	exportMessage = 'The Roadmap export has halted\n' + 'Time halted: ' + timeUpdated
	MessageBox.Show("The Roadmap export has halted", "OSC RoadMap Export Prompt", MessageBoxButtons.OK)

Document.Properties['exportMessage'] = exportMessage
print(exportMessage)
