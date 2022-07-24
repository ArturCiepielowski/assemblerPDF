Sub Main()

Dim openDoc As AssemblyDocument
openDoc = ThisDoc.Document
Dim oDoc As Document


ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
Dim count As Integer = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count



Call creatingPDF()

Call firstLoop(openDoc, oDoc, count)

End Sub





Sub firstLoop (openDoc As AssemblyDocument, oDoc As Document, count As Integer)

	For item As Integer = 1 To count ' ----------------- Main FOR
		On Error Resume Next

		pathMap = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Document.DisplayName

		MsgBox(pathMap)
		
		
		fullName=openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Document.FullDocumentName
		assemblyType = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Type
		Dim oPart As AssemblyDocument
		
		if assemblyType <>	kAssemblyComponentDefinitionObject Then 
			
			'ThisApplication.ActiveDocument.Close(True)
			Exit For
		
		Else if assemblyType = kAssemblyComponentDefinitionObject
			oPart = ThisApplication.Documents.Open(fullName, True)

			ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
			ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
			newCount = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
			
			Call creatingPDF()
			Call secondLoop(newCount)
			
		End if
Next

End Sub





Sub secondLoop (newCount)

For newItem As Integer = 1 To newCount ' ---------------[0] FOR 
		
			
	On Error Resume Next
			
	pathMap0 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.DisplayName

	MsgBox(pathMap0)
			
	newFullName=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.FullDocumentName
	newAssemblyType = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Type
	Dim NewoPart As AssemblyDocument
	Dim closePart As PartDocument
			
			
	if newAssemblyType = kAssemblyComponentDefinitionObject
		NewoPart = ThisApplication.Documents.Open(newFullName, True)
		
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
		
		Call creatingPDF()		
		Call thirdLoop(newCount1)
				
	Else if newAssemblyType <>	kAssemblyComponentDefinitionObject Then 
		ThisApplication.ActiveDocument.Close(True)
		Exit For
				
			End if
		Next
End Sub






Sub thirdLoop (newCount1)
For newItem1 As Integer = 1 To newCount1 ' ----------------[1] FOR
				
						
On Error Resume Next
					
					
	pathMap1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.DisplayName

	MsgBox(pathMap1)
					
	newFullName1=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem1).ComponentDefinitions.Item(1).Document.FullDocumentName
	newAssemblyType1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem1).ComponentDefinitions.Item(1).Type
	Dim NewoPart1 As AssemblyDocument
	Dim closePart1 As PartDocument
					
	if newAssemblyType1 = kAssemblyComponentDefinitionObject
		NewoPart1 = ThisApplication.Documents.Open(newFullName1, True)
			
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount2 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
		
		Call creatingPDF()
		Call fourthLoop(newCount2)
		
	Else if newAssemblyType1 <>	kAssemblyComponentDefinitionObject Then 
		ThisApplication.ActiveDocument.Close(True)
		Exit For
						
	End if
Next

End Sub




Sub fourthLoop(newCount2)

For newItem2 As Integer = 1 To newCount2 ' ----------------[2] FOR
							
	On Error Resume Next
						
							
	pathMap2 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.DisplayName

	MsgBox(pathMap2)
	newFullName2=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem2).ComponentDefinitions.Item(1).Document.FullDocumentName
	newAssemblyType2 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem2).ComponentDefinitions.Item(1).Type
	Dim NewoPart2 As AssemblyDocument
	Dim closePart2 As PartDocument
							
					
	if newAssemblyType2 = kAssemblyComponentDefinitionObject
		NewoPart2 = ThisApplication.Documents.Open(newFullName2, True)
		Call creatingPDF()
						
	Else if newAssemblyType2 <>	kAssemblyComponentDefinitionObject Then 
						
		ThisApplication.ActiveDocument.Close(True)
		Exit For
								
	End if
Next

End Sub





Sub creatingPDF()

Dim oDoc As Document
oDoc = ThisApplication.ActiveDocument

	
	Dim oDocName As String = oDoc.FullFileName
	Dim oDocJustName As String = oDoc.DisplayName
	
	
	Dim sFileName As String = Split(oDocName, oDocJustName)(0)
	
	Dim  displayNameCut As String = Split(oDocJustName, ".iam")(0)
	
	Dim sDrawingName As String = sFileName & "_RYSUNKI\_WYKONAWCZE\" & displayNameCut & ".idw"
	

ThisApplication.Documents.Open(sDrawingName, True)

Dim drwoDoc As Document
drwoDoc = ThisApplication.ActiveDocument

Call MakePDFFromDoc(drwoDoc)

ThisApplication.ActiveDocument.Close

End Sub

Sub MakePDFFromDoc(ByRef docDrw As Document)
 

	Dim docPath As String = docDrw.FullFileName
	Dim  docBlank As String = Split(docPath, ".idw")(0)
	
	
	
  oPDFAddIn = ThisApplication.ApplicationAddIns.ItemById _
  ("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
  oContext = ThisApplication.TransientObjects.CreateTranslationContext
  oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
  oOptions = ThisApplication.TransientObjects.CreateNameValueMap
  oDataMedium = ThisApplication.TransientObjects.CreateDataMedium
  
  


  oOptions.Value("All_Color_AS_Black") = 0
  oOptions.Value("Remove_Line_Weights") = 0
  oOptions.Value("Vector_Resolution") = 400
  oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
 

  oDataMedium.FileName = docBlank  & ".pdf"

 
   oPDFAddIn.SaveCopyAs(docDrw, oContext, oOptions, oDataMedium)'
  
  
End Sub


