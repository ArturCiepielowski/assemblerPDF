Dim openDoc As AssemblyDocument
openDoc = ThisDoc.Document
Dim oDoc As Document
Dim counter As Integer
counter = 0

ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
count = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count

For item As Integer = 1 To count ' ----------------- Main FOR
	On Error Resume Next
	
	
	fullName=openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Document.FullDocumentName
	assemblyType = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Type
	Dim oPart As PartDocument
	
	if assemblyType <>	kAssemblyComponentDefinitionObject Then 
		'MsgBox("To nie jest zlozenie no2")
		Exit For
	
	Else if assemblyType = kAssemblyComponentDefinitionObject
		oPart = ThisApplication.Documents.Open(fullName, True)
		counter =counter +1
		
		
		
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
		For newItem As Integer = 1 To newCount ' ---------------[0] FOR 
		
			MsgBox(newCount)	
			On Error Resume Next
			newFullName=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.FullDocumentName
			newAssemblyType = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Type
			Dim NewoPart As PartDocument
			
			if newAssemblyType <>	kAssemblyComponentDefinitionObject Then 
				'MsgBox("To nie jest zlozenie no1")
				Exit For
			Else if newAssemblyType = kAssemblyComponentDefinitionObject
				NewoPart = ThisApplication.Documents.Open(newFullName, True)
				counter =counter +1
				
				
				
				ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
				ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
				newCount1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
			
				For newItem1 As Integer = 1 To newCount1 ' ----------------[1] FOR
				
					MsgBox(newCount1)	
					On Error Resume Next
					newFullName1=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem1).ComponentDefinitions.Item(1).Document.FullDocumentName
					newAssemblyType1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem1).ComponentDefinitions.Item(1).Type
					Dim NewoPart1 As PartDocument
					
					if newAssemblyType1 <>	kAssemblyComponentDefinitionObject Then 
						'MsgBox("To nie jest zlozenie no1")
						Exit For
					Else if newAssemblyType1 = kAssemblyComponentDefinitionObject
						NewoPart1 = ThisApplication.Documents.Open(newFullName1, True)
						counter1 =counter +1
						
						
					
					End if
			Next
			
			
			End if
		Next
			
			
	
	End if
Next


	
	
	' ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
	' ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
	' newCount = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
	' For newItem As Integer = 1 To newCount
		' On Error Resume Next
		' newFullName=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.FullDocumentName
		' newAssemblyType = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Type
		' Dim NewoPart As PartDocument
		' if assemblyType = kAssemblyComponentDefinitionObject
			' NewoPart = ThisApplication.Documents.Open(newFullName, True)
			' counter =counter +1
			
		' Else if newAssemblyType <>	kAssemblyComponentDefinitionObject Then 
			' 'MsgBox("To nie jest zlozenie")
			' Exit For
		' End if
	' Next


'MsgBox(counter)

