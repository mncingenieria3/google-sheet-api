Attribute VB_Name = "Módulo1"
Sub CallAPI(params As String)
    ' Declare variables
    Dim objHTTP As Object
    Dim URL As String
    Dim RequestBody As String
    Dim SourceBook As Workbook
    Dim OriginWorkSheet As Worksheet
    Dim proveedor As String
    Dim fecha As Date
    Dim valorBruto As Double
    Dim valor As Double
    Dim numeroCot As String
    Dim observaciones As String
        
    ' IDK, but, I'm not going to refactor this function
    Set SourceBook = ActiveWorkbook
    Set OriginWorkSheet = SourceBook.Sheets(params)
    proveedor = OriginWorkSheet.Range("B8").Value
    fecha = OriginWorkSheet.Range("B5").Value
    valorBruto = OriginWorkSheet.Range("H48").Value
    valor = OriginWorkSheet.Range("H52").Value
    solicitante = OriginWorkSheet.Range("F5").Value
    centroCosto = OriginWorkSheet.Range("H5").Value
    ciudad = OriginWorkSheet.Range("H6").Value
    numeroCot = OriginWorkSheet.Range("G10").Value
    observaciones = OriginWorkSheet.Range("A69").Value
    
    ' Create a WinHttpRequest object
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Set the URL you want to send the POST request to
    URL = OriginWorkSheet.Range("B3").Value
    MsgBox "API URL: " & URL

    ' Set the request body (raw JSON or other content)
    RequestBody = "[{""PROVEEDOR"": """ & proveedor & """, ""FECHA"": """ & fecha & """, ""VALOR_BRUTO"": """ & valorBruto & """, ""VALOR"": """ & valor & """, ""N_COT"": """ & numeroCot & """, ""SOLICITANTE"": """ & solicitante & """, ""CENTRO_COSTO"": """ & centroCosto & """, ""CIUDAD"": """ & ciudad & """, ""OBSERVACIONES"": """ & observaciones & """}]"

    ' Open a connection to the URL
    objHTTP.Open "POST", URL, False

    ' Set request headers (optional)
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.setRequestHeader "Content-Type", "application/json" ' Set the appropriate content type

    ' Send the request with the request body
    objHTTP.send RequestBody

    ' Check if the request was successful (status code 200)
    If objHTTP.Status = 200 Then
        ' Print the response text to the Immediate window or do something with it
        responseText = objHTTP.responseText
        MsgBox "Consecutivo asignado: " & responseText, vbInformation
        OriginWorkSheet.Range("H2").Value = responseText
        Debug.Print objHTTP.responseText
    Else
        ' Handle errors or display the status code and response text
        Debug.Print "Request failed. Status code: " & objHTTP.Status & vbCrLf & "Response: " & objHTTP.responseText
    End If

    ' Clean up the WinHttpRequest object
    Set objHTTP = Nothing
End Sub
Sub ExportOC(params As String)
    ' Declare variables
    Dim ws As Worksheet
    Dim newWB As Workbook
    Dim fileName As String
    Dim filePath As String
    Dim OC_Num As String
    Dim documentType As String

    ' Define the sheet you want to export
    Set ws = ThisWorkbook.Sheets(params)
    
    ' Get the current year and OC number
    currentYear = Year(Date)
    OC_Num = ws.Range("G2").Value

    ' Create the file name based on your pattern
    If params = "ORDEN DE COMPRA" Then
        documentType = "OC"
    Else
        documentType = "OS"
    End If
    fileName = "MNC-" & documentType & "-" & currentYear & "-" & OC_Num

    ' Prompt the user to choose the file path
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder to Save New Workbook"
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub ' User canceled, so exit the subroutine
        End If
    End With

    ' Check if the file already exists and delete it if necessary
    If Dir(filePath & "\" & fileName) <> "" Then
        Kill filePath & "\" & fileName
    End If

    ' Copy the sheet's values to a new workbook
    ws.Copy
    Set newWB = ActiveWorkbook

    ' Paste values only to remove formulas
    With newWB.Sheets(1).UsedRange
        .Value = .Value
    End With
    
    ' Export the sheet as PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=filePath & "\" & fileName & ".pdf"
    
    ' Save the new workbook with the specified file name and path
    newWB.SaveAs filePath & "\" & fileName & ".xlsx"
    newWB.Close SaveChanges:=False

    ' Clean up
    Set newWB = Nothing
    Set ws = Nothing
    Sheets("REQUISICION_OS").Select
    Range("B1").Select
    MsgBox documentType & " guardada en " & filePath, vbInformation
End Sub
Sub CheckAndWakeupAPI(params As String)
    ' Declare variables
    Dim URL As String
    Dim XMLHttpRequest As Object
    Dim SourceBook As Workbook
    Dim OriginWorkSheet As Worksheet
    
    Set SourceBook = ActiveWorkbook
    Set OriginWorkSheet = SourceBook.Sheets(params)
    URL = OriginWorkSheet.Range("B3").Value
    MsgBox "API URL: " & URL
    

    ' Create a new XMLHttpRequest object
    Set XMLHttpRequest = CreateObject("MSXML2.ServerXMLHTTP")

    ' Open a connection to the URL
    XMLHttpRequest.Open "GET", URL, False

    ' Send the GET request
    XMLHttpRequest.send

    ' Check if the request was successful (status code 200)
    If XMLHttpRequest.Status = 200 Then
        ' Print the response text to the Immediate window or do something with it
        MsgBox "API URL: " & "Funcionando"
    Else
        ' Handle errors or display the status code and response text
        MsgBox "API URL: " & "Ha ocurrido un error con la URL", vbExclamation
    End If

    ' Clean up the XMLHttpRequest object
    Set XMLHttpRequest = Nothing
End Sub
Sub ExportSheetToNewWorkbookNoFormulas()
    Dim ws As Worksheet
    Dim newWB As Workbook
    Dim fileName As String
    Dim filePath As String
    Dim currentYear As String

    ' Define the sheet you want to export
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet's name

    ' Get the current year
    currentYear = Year(Date)

    ' Create the file name based on your pattern
    fileName = "MNC-OC_" & currentYear & "01.xlsx"

    ' Prompt the user to choose the file path
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder to Save New Workbook"
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub ' User canceled, so exit the subroutine
        End If
    End With

    ' Check if the file already exists and delete it if necessary
    If Dir(filePath & "\" & fileName) <> "" Then
        Kill filePath & "\" & fileName
    End If

    ' Copy the sheet's values to a new workbook
    ws.Copy
    Set newWB = ActiveWorkbook

    ' Paste values only to remove formulas
    With newWB.Sheets(1).UsedRange
        .Value = .Value
    End With

    ' Save the new workbook with the specified file name and path
    newWB.SaveAs filePath & "\" & fileName
    newWB.Close SaveChanges:=False

    ' Clean up
    Set newWB = Nothing
    Set ws = Nothing

    MsgBox "Sheet exported to " & filePath & "\" & fileName & " with values only (no formulas)", vbInformation
End Sub
