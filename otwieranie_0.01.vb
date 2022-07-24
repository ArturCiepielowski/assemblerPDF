Dim openDoc As AssemblyDocument
openDoc = ThisDoc.Document
Dim oDoc As Document
Dim counter As Integer
counter = 0


ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
count = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
For item As Integer = 1 To count
	On Error Resume Next
	fullName=openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Document.FullDocumentName
	assemblyType = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Type
	Dim oPart As PartDocument
	if assemblyType = kAssemblyComponentDefinitionObject
		oPart = ThisApplication.Documents.Open(fullName, True)
		counter =counter +1
		
	Else if assemblyType <>	kAssemblyComponentDefinitionObject Then 
		'MsgBox("To nie jest zlozenie")
		Exit For
	End if
	
	
	ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
	ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
	newCount = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
	For newItem As Integer = 1 To newCount
		newFullName=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.FullDocumentName
		newAssemblyType = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Type
		Dim NewoPart As PartDocument
	if assemblyType = kAssemblyComponentDefinitionObject
		NewoPart = ThisApplication.Documents.Open(newFullName, True)
		counter =counter +1
		
	Else if newAssemblyType <>	kAssemblyComponentDefinitionObject Then 
		'MsgBox("To nie jest zlozenie")
		Exit For
	End if
	Next
Next

MsgBox(counter)