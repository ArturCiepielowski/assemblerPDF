Dim openDoc As AssemblyDocument
openDoc = ThisDoc.Document
Dim oDoc As Document


ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
count = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count

For item As Integer = 1 To count ' ----------------- Main FOR
	On Error Resume Next
	
	MsgBox(count)
	
	fullName=openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Document.FullDocumentName
	assemblyType = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Type
	Dim oPart As AssemblyDocument
	
	if assemblyType <>	kAssemblyComponentDefinitionObject Then 
		
		
		Exit For
	
	Else if assemblyType = kAssemblyComponentDefinitionObject
		oPart = ThisApplication.Documents.Open(fullName, True)
		
		
		
		
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
		For newItem As Integer = 1 To newCount ' ---------------[0] FOR 
		
			
			On Error Resume Next
			MsgBox(newCount)	
			newFullName=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.FullDocumentName
			newAssemblyType = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Type
			Dim NewoPart As AssemblyDocument
			Dim closePart As PartDocument
			'MsgBox(newFullName)
			'MsgBox(newAssemblyType)
			
			if newAssemblyType = kAssemblyComponentDefinitionObject
				NewoPart = ThisApplication.Documents.Open(newFullName, True)
				
				
				
				
				ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
				ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
				newCount1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
			
				For newItem1 As Integer = 1 To newCount1 ' ----------------[1] FOR
				
						
					On Error Resume Next
					'MsgBox(newCount1)
					newFullName1=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem1).ComponentDefinitions.Item(1).Document.FullDocumentName
					newAssemblyType1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem1).ComponentDefinitions.Item(1).Type
					Dim NewoPart1 As AssemblyDocument
					Dim closePart1 As PartDocument
					'MsgBox(newFullName1)
					'MsgBox(newAssemblyType1)
					
					if newAssemblyType1 = kAssemblyComponentDefinitionObject
						NewoPart1 = ThisApplication.Documents.Open(newFullName1, True)
						
						
						ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
						ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
						newCount2 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
						
						For newItem2 As Integer = 1 To newCount2 ' ----------------[2] FOR
							
							On Error Resume Next
							'MsgBox(newCount2)
							newFullName2=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem2).ComponentDefinitions.Item(1).Document.FullDocumentName
							newAssemblyType2 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem2).ComponentDefinitions.Item(1).Type
							Dim NewoPart2 As AssemblyDocument
							Dim closePart2 As PartDocument
							'MsgBox(newFullName2)
							'MsgBox(newAssemblyType2)
					
							if newAssemblyType2 = kAssemblyComponentDefinitionObject
								NewoPart2 = ThisApplication.Documents.Open(newFullName2, True)
						
						
							Else if newAssemblyType2 <>	kAssemblyComponentDefinitionObject Then 
						
								ThisApplication.ActiveDocument.Close(True)
								Exit For
								
							End if
						Next
					
					Else if newAssemblyType1 <>	kAssemblyComponentDefinitionObject Then 
						ThisApplication.ActiveDocument.Close(True)
						Exit For
						
					End if
				Next
			Else if newAssemblyType <>	kAssemblyComponentDefinitionObject Then 
				
				Exit For
				oPart.Close
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

