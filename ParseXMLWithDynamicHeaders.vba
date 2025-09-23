Sub ParseXMLWithDynamicHeaders()
  Dim http As Object
  'XML
  Dim xmlDoc As Object
  Dim schemaNode As Object
  Dim elementNodes As Object
  Dim rowNodes As Object
  Dim rowNode As Object
  Dim nodeBook As Object
  Dim attributeID As Object
  Dim attributeName As Object

  Dim headerName As String
  '
  Dim ws As Worksheet
  '
  Dim rows As Object
  Dim row As Object
  Dim childnode As Object
  Dim ChildNodesCounter As Integer
  Dim cellrow As Integer, cnt As Integer
  Dim columncount As Integer
  Dim colunnames() As String
  Dim arraylength As Integer

 
  With ThisWorkbook.Sheets(1)
    ' Set the target worksheet and clean contents before starting
    Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.ClearContents
    
    ' Get HTTP request
    ' Define the URL
    url = "https://{YOUR ANALYTICS URL}?path={PATH TO THE REPORT}&limit=25&col_names=true&apikey={YOUR API KEY}"
    ' Create HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send

    If http.Status <> 200 Then
        MsgBox "Failed to download XML. Status: " & http.Status
        Exit Sub
    End If

    ' Load XML from response
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.LoadXML http.responseText

    If xmlDoc.ParseError.ErrorCode <> 0 Then
        MsgBox "Error in XML file: " & xmlDoc.ParseError.Reason
        Exit Sub
    End If

    ' Get Column Headers
    Set schemaNode = xmlDoc.SelectSingleNode("//*[local-name()='schema']/*[local-name()='complexType']/*[local-name()='sequence']")
    Set elementNodes = schemaNode.SelectNodes("*[local-name()='element']")
    ' Add columns headers to the worksheet
    ' and save the column names in array colunnames()
    arraylength = elementNodes.length
    ReDim Preserve colunnames(arraylength)
    For Each nodeBook In elementNodes
        If Not attributeID Is Nothing Then
            'columnHeading is found in the node
            Set attributeID = nodeBook.Attributes.getNamedItem("saw-sql:columnHeading")
        Else
            'If columnHeading is not found, use name instead
            Set attributeID = nodeBook.Attributes.getNamedItem("name")
        End If
        ' Fill column headings in the worksheet
        ws.Cells(1, cnt + 1).Value = attributeID.NodeValue
        ' Add column headings to colunnames()
        colunnames(cnt) = nodeBook.Attributes.getNamedItem("name").NodeValue
        cnt = cnt + 1
    Next nodeBook
    
    ' Fill rows
    ' Extract <Row> elements (namespace-aware)
    Set rows = xmlDoc.SelectNodes("//*[local-name()='Row']")
    
    If rows Is Nothing Or rows.length = 0 Then
        MsgBox "No <Row> elements found."
        Exit Sub
    End If
    
    ' Loop through rows and write data starting from the 2nd row
    cellrow = 2
    For Each row In rows
        For ChildNodesCounter = 0 To row.ChildNodes.length - 1
            ' Find the correct column for the value
            ' by matching the column name with childe node name
            i = 0
            For i = 0 To UBound(colunnames) - 1
              If colunnames(i) = row.ChildNodes(ChildNodesCounter).NodeName Then
                ' Insert the value in the correct column
                ws.Cells(cellrow, i + 1).Value = row.ChildNodes(ChildNodesCounter).Text
              End If
            Next i
            
        Next ChildNodesCounter
        cellrow = cellrow + 1
    Next row


  End With

End Sub


