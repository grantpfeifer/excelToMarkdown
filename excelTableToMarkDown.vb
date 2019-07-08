Sub rangeToMarkDown()

    Dim cell As Range
    Dim selectedRange As Range

Set selectedRange = Application.Selection

Dim rowCounter As Integer
    Dim columnCounter As Integer
    Dim totalColumns As Integer
    Dim currentColumnWidth As Integer
    Dim linkLength As Integer
    Dim s As String
    Dim n As Integer
    Dim userName As String
    Dim filePath As String

    Dim fileName As Variant
    fileName = InputBox("Name for Markdown File (do not include file extension):")


    'get username and create export file on desktop
    userName = Environ("Username")
    filePath = "C:\users\" + userName + "\Desktop\" + fileName + ".txt"

    n = FreeFile()
    Open filePath For Output As #n

totalColumns = selectedRange.Columns.Count

    Dim columnWidth(50) As String    'maximum of 50 columns


    'init lengths of columns
    For I = 0 To totalColumns
        columnWidth(I) = 0
    Next I


    'go through range to calculate maximum lengths of each column
    For Each Row In selectedRange.Rows

        columnCounter = 0

        For Each cell In Row.Cells

            On Error Resume Next
            linkAddress = vbNullString
            linkAddress = cell.Hyperlinks(1).Address
            On Error GoTo 0

            If linkAddress = vbNullString Then
                currentColumnWidth = Len(cell.Value)
            Else
                temp = linkAddress
                linkAddress = Replace(temp, " ", "%20")
                linkLength = Len("[" & cell.Value & "]" & "(" & linkAddress & ")")
                currentColumnWidth = linkLength
            End If


            If (currentColumnWidth > columnWidth(columnCounter)) Then

                columnWidth(columnCounter) = currentColumnWidth

            End If

            columnCounter = columnCounter + 1

        Next cell

    Next Row


    '/// go through range to calculate maximum lengths of each column
    Dim currentLine As String

    rowCounter = 0
    For Each Row In selectedRange.Rows

        columnCounter = 0

        currentLine = "|"

        For Each cell In Row.Cells

            currentColumnWidth = columnWidth(columnCounter)
            Dim extraSpaces As Integer


            currentLine = currentLine & " "

            On Error Resume Next
            linkAddress = vbNullString
            linkAddress = cell.Hyperlinks(1).Address
            On Error GoTo 0

            If linkAddress = vbNullString Then
                currentLine = currentLine & cell.Value
                extraSpaces = currentColumnWidth - Len(cell.Value)
            Else
                temp = linkAddress
                linkAddress = Replace(temp, " ", "%20")
                currentLine = currentLine & "[" & cell.Value & "]" & "(" & linkAddress & ")"
                extraSpaces = currentColumnWidth - Len("[" & cell.Value & "]" & "(" & linkAddress & ")")
            End If


            For j = 0 To extraSpaces

                currentLine = currentLine & " "


            Next j

            If (columnCounter <> totalColumns - 1) Then
                currentLine = currentLine & " |"
            End If

            columnCounter = columnCounter + 1

        Next cell

        Debug.Print currentLine
    Print #n, currentLine

    If (rowCounter = 0) Then

            currentLine = "|"
            columnCounter = 0

            For j = 0 To (totalColumns - 2)

                currentLine = currentLine
                currentColumnWidth = columnWidth(columnCounter)
                currentLine = currentLine & "-"

                For k = 0 To currentColumnWidth

                    currentLine = currentLine & "-"
                Next k

                currentLine = currentLine & "-|"
                columnCounter = columnCounter + 1

            Next j
            currentColumnWidth = columnWidth(columnCounter)
            currentLine = currentLine & "-"

            For k = 0 To currentColumnWidth

                currentLine = currentLine & "-"
            Next k

            currentLine = currentLine & "-"
            columnCounter = columnCounter + 1


            Debug.Print currentLine
        Print #n, currentLine
    End If

        rowCounter = rowCounter + 1

    Next Row

    Close #n

MsgBox "File saved to " + filePath

End Sub