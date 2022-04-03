Installing the Addon:

	* To install the Addon go to your local Citavi install directory and navigate to the .\Citavi\AddOns Folder. 
	* Copy the jsonExportAddon.dll into the AddOns Folder.
	* Restart your Citavi and the Addon will be loaded.

Using the Addon:
	
	* After installing the Addon open your local Citavi and the desired Project
	* In the top Toolbar under "Tools" select "Export Project to JSON"
	* A promp will ask you to select a desired location where the export will be saved.
	* After choosing a location the export starts and the following folder/file structure will be created:

    +---CitaviExport_$ProjectName$\
	|
	|
	+--- categories.json
	+--- keywords.json
	+--- knowledgeItems.json
	+--- references.json

	+---Files\
	|   |
	|   +--- OrigFilename or FileReferenceID

	+---log.txt

	
	NOTE:   The export can take a while espacially for large projects with alot of references that need to be downloaded. The Citavi program is unresponsive in this time
		After the export is finished a popup will inform the user.
		The log.txt files contains information about the downloaded and extracted pdf files and will contain any errors that occured during export.

