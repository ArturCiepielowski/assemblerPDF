Dim openDoc As AssemblyDocument
openDoc = ThisDoc.Document
Dim oDoc As Document
Dim print As String


count = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
For item As Integer = 1 To count
	On Error Resume Next
	print=openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Document.FullDocumentName
	assemblyType = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Type
	Dim oPart As PartDocument
		if assemblyType = kAssemblyComponentDefinitionObject
		oPart = ThisApplication.Documents.Open(print, True)
		End if

Next
