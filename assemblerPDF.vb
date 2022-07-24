Imports System.IO
Imports System.Text

Sub Main()

Dim openDoc As AssemblyDocument
openDoc = ThisDoc.Document
Dim oDoc As Document = ThisApplication.ActiveDocument
Dim oDocNameMain As String = oDoc.FullFileName

Dim mainPath As String = Split(oDocNameMain, oDoc.DisplayName)(0)




ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
Dim count As Integer = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count

Dim counter As Integer = 1

'MsgBox(counter & "- Glowne zlozenie")

creatingPDF(counter, mainPath)

firstLoop(openDoc, oDoc, count, counter, mainPath)

End Sub


'-------------------------------------------------------------------------- pierwsza petla -----------------------------------------------------------------


Function firstLoop (openDoc As AssemblyDocument, oDoc As Document, count As Integer, counter As Integer, mainPath As String)

Dim check as Integer = 0
Dim breaker As Integer = 0

For item As Integer = 1 To count 
	On Error Resume Next

	counter= counter + 1
	
	check=check+1
	'pathMap = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Document.DisplayName

	'MsgBox(counter &  "- Zlozenie I poz."& " Ilosc " & item & " Z : "& count)
		
		
	fullName=openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Document.FullDocumentName
	assemblyType = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Type
	Dim oPart As AssemblyDocument
	
		
	if assemblyType <>	kAssemblyComponentDefinitionObject Then 
			
		'ThisApplication.ActiveDocument.Close(True)
		braker=1
		'MsgBox("Wyjscie: to nie jest złożenie")
		Exit For
		
	Else if assemblyType = kAssemblyComponentDefinitionObject
		oPart = ThisApplication.Documents.Open(fullName, True)

		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
			
		creatingPDF(counter, mainPath)
		counter = secondLoop(newCount, counter, mainPath)
			
	End if
Next

'MsgBox(check &" : " & count)

If check = count And breaker=0

ThisApplication.ActiveDocument.Close(True)
'MsgBox("Koniec elementów")

End if

return counter

End Function


'-------------------------------------------------------------------------- druga petla -----------------------------------------------------------------


Function secondLoop (newCount, counter, mainPath)

Dim newCheck as Integer = 0
Dim NewBreaker As Integer = 0

For newItem As Integer = 1 To newCount 
			
	On Error Resume Next
	
	
	
	newCheck = newCheck+1
	
	counter = counter + 1
			
	'pathMap0 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.DisplayName

	'MsgBox(counter & "- Zlozenie II poz."& " Ilosc " & newItem & " Z : "& newCount )
			
	newFullName=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.FullDocumentName
	newAssemblyType = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Type
	Dim NewoPart As AssemblyDocument
	Dim closePart As PartDocument
	
			
			
	if newAssemblyType = kAssemblyComponentDefinitionObject
		NewoPart = ThisApplication.Documents.Open(newFullName, True)
		
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
		
		creatingPDF(counter, mainPath)		
		counter = thirdLoop(newCount1, counter, mainPath)
				
	Else if newAssemblyType <>	kAssemblyComponentDefinitionObject Then 
		
		NewBreaker = 1
		ThisApplication.ActiveDocument.Close(True)
		
		'MsgBox("Wyjscie: to nie jest złożenie")
		Exit For
				
			End if
		Next
		
		
		'MsgBox(newCheck &" : " & newCount)

		If newCheck = newCount And NewBreaker=0

			ThisApplication.ActiveDocument.Close(True)
			'MsgBox("Koniec elementów")

		End if
		
		
		return counter
		
End Function



'-------------------------------------------------------------------------- trzecia petla -----------------------------------------------------------------


Function thirdLoop (newCount1, counter, mainPath)

Dim newCheck1 as Integer = 0
Dim NewBreaker1 As Integer = 0

For newItem1 As Integer = 1 To newCount1 
				
						
On Error Resume Next

	newCheck1 = newCheck1+1
	
	counter= counter + 1	
	
	'pathMap1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.DisplayName

	'MsgBox(counter & "- Zlozenie III poz."& " Ilosc " & newItem1 & " Z : "& newCount1)
					
	newFullName1=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem1).ComponentDefinitions.Item(1).Document.FullDocumentName
	newAssemblyType1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem1).ComponentDefinitions.Item(1).Type
	Dim NewoPart1 As AssemblyDocument
	Dim closePart1 As PartDocument
	'MsgBox(newAssemblyType1)
					
	if newAssemblyType1 = kAssemblyComponentDefinitionObject
		NewoPart1 = ThisApplication.Documents.Open(newFullName1, True)
			
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount2 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
		
		creatingPDF(counter, mainPath)
		counter =fourthLoop(newCount2, counter, mainPath)
		
	Else if newAssemblyType1 <>	kAssemblyComponentDefinitionObject Then 
		
		NewBreaker1 = 1
		ThisApplication.ActiveDocument.Close(True)
		'MsgBox("Wyjscie: to nie jest złożenie")
		Exit For
						
	End if
Next

'MsgBox(newCheck1 &" : " & newCount1)

		If newCheck1 = newCount1 And NewBreaker1 = 0

	ThisApplication.ActiveDocument.Close(True)
	'MsgBox("Koniec elementów")

End if
		
return counter

End Function


'-------------------------------------------------------------------------- czwarta petla -----------------------------------------------------------------

Function fourthLoop (newCount2, counter, mainPath)

Dim newCheck2 as Integer = 0
Dim NewBreaker2 As Integer=0

For newItem2 As Integer = 1 To newCount2 
				
						
On Error Resume Next
	
	
	
	newCheck2 = newCheck2+1
		
	counter= counter + 1	
	
	'pathMap1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.DisplayName

	'MsgBox(counter & "- Zlozenie IV poz."& " Ilosc " & newItem2 & " Z : "& newCount2)
					
	newFullName2=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem2).ComponentDefinitions.Item(1).Document.FullDocumentName
	newAssemblyType2 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem2).ComponentDefinitions.Item(1).Type
	
	
	Dim NewoPart2 As AssemblyDocument
	Dim closePart2 As PartDocument
	
	'MsgBox(newAssemblyType2)				
	if newAssemblyType2 = kAssemblyComponentDefinitionObject
		NewoPart2 = ThisApplication.Documents.Open(newFullName2, True)
			
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount3 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
		
		creatingPDF(counter, mainPath)
		counter = fifthLoop(newCount3, counter, mainPath)
		
	Else if newAssemblyType2 <>	kAssemblyComponentDefinitionObject Then 
		NewBreaker2=1
		ThisApplication.ActiveDocument.Close(True)
		'MsgBox("Wyjscie: to nie jest złożenie")
		Exit For
						
	End if
Next

'MsgBox(newCheck2 &" : " & newCount2)

If newCheck2 = newCount2 And NewBreaker2 =0

	ThisApplication.ActiveDocument.Close(True)
	'MsgBox("Koniec elementów")

End if

return counter

End Function


'-------------------------------------------------------------------------- piąta petla -----------------------------------------------------------------

Function fifthLoop(newCount3, counter, mainPath)

Dim newCheck3 as Integer = 0
Dim NewBreaker3 As Integer = 0

For newItem3 As Integer = 1 To newCount3 
							
	On Error Resume Next
	
	
	
	newCheck3 = newCheck3 + 1
	
	counter= counter + 1					
							
	'pathMap2 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.DisplayName

	'MsgBox(counter & "- Zlozenie V poz."& " Ilosc " & newItem3 & " Z : "& newCount3)
	
	newFullName3=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem3).ComponentDefinitions.Item(1).Document.FullDocumentName
	newAssemblyType3 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem3).ComponentDefinitions.Item(1).Type
	Dim NewoPart3 As AssemblyDocument
	Dim closePart3 As PartDocument
	
							
					
	if newAssemblyType3 = kAssemblyComponentDefinitionObject
		NewoPart3 = ThisApplication.Documents.Open(newFullName3, True)
		creatingPDF(counter, mainPath)
						
	Else if newAssemblyType3 <>	kAssemblyComponentDefinitionObject Then 
		
		NewBreaker3=1
		ThisApplication.ActiveDocument.Close(True)
		'MsgBox("Wyjscie: to nie jest złożenie")
		Exit For
								
	End if
Next

'MsgBox(newCheck3 &" : " & newCount3)

If newCheck3 = newCount3 And NewBreaker3=0

	ThisApplication.ActiveDocument.Close(True)
	'MsgBox("Koniec elementów")

End if
return counter

End Function





'-------------------------------------------------------------------------- funkcja tworząca PDF -----------------------------------------------------------------


Function creatingPDF(counter, mainPath)

Dim pdfCounter As Integer= counter 
Dim oDoc As Document
oDoc = ThisApplication.ActiveDocument

	
	Dim oDocName As String = oDoc.FullFileName
	Dim oDocJustName As String = oDoc.DisplayName
	
	
	Dim sFileName As String = Split(oDocName, oDocJustName)(0)
	
	'MsgBox(sFileName)
	'MsgBox(mainPath)
	
	'If sFileName = mainPath
	
	'	MsgBox("Plik ma tą samą lokalizację")
		
	'Else	
	
	'	MsgBox("Pliki mają różne lokalizację")
		
	'End If
	
	Dim  displayNameCut As String = Split(oDocJustName, ".iam")(0)
	
	Dim sDrawingName As String = sFileName & "_RYSUNKI\WYKONAWCZE\" & displayNameCut & ".idw"
	
Try	
	If Not System.IO.File.Exists(sDrawingName) Then
		'MsgBox("Nie ma rysunku")
		
		errorLog (sDrawingName, sFileName)
		
	End If
Catch
End Try	
ThisApplication.Documents.Open(sDrawingName, True)


Dim drwoDoc As Document
drwoDoc = ThisApplication.ActiveDocument

MakePDFFromDoc(drwoDoc, sFileName, pdfCounter)

ThisApplication.ActiveDocument.Close

End Function

Function MakePDFFromDoc(docDrw, sFileName, counter)
 

	Dim docJustName As String = docDrw.DisplayName
	Dim docNameCut As String = Split(docJustName, ".idw")(0)

	Dim newPDFname As String = counter & "." & docNameCut
	
	
	
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
 

  oDataMedium.FileName =  sFileName & "_RYSUNKI\WYKONAWCZE\" & newPDFname  & ".pdf"

  oPDFAddIn.SaveCopyAs(docDrw, oContext, oOptions, oDataMedium)
  
  
End Function

Function errorLog (sDrawingName, sFileName)

Dim myDate As String = Now().ToString("yyyy-MM-dd HH.m.ss")
myDate = myDate.Replace(":","")  

Dim path As String = sFileName & "_RYSUNKI\WYKONAWCZE\" &"drwLog.txt"

Dim file As System.IO.StreamWriter
file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
	
file.WriteLine(sDrawingName)

file.Close()
End Function





