Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO

Module myFunc

    '===================================================================================
    '             === Load database ===
    '===================================================================================

    Sub loadDataBaseFolder()


        '   1.   Get database folder over Folder browser

        mainForm.FBD.SelectedPath = Directory.GetCurrentDirectory()
        If (mainForm.FBD.ShowDialog() = DialogResult.OK) Then
            mainForm.sDir = mainForm.FBD.SelectedPath
        End If



        '   2.   Get list of database files , names of each file, list of worksheets in each file

        Dim name As String                           ' variable to get name of database file

        mainForm.fileNames = New Collection         ' collection of names of each file

        Dim key As Integer = 0

        mainForm.i_superPivotDict = New Dictionary(Of Integer, Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable)))
        mainForm.i_pivotTableDict = New Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable))
        mainForm.i_pivot_wsDict = New Dictionary(Of Integer, Dictionary(Of Integer, ExcelWorksheet))

        For Each foundFile As String In My.Computer.FileSystem.GetFiles _
        (mainForm.sDir, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.xlsx")

            '   Extract file name from full path

            mainForm.sFilePath = CStr(foundFile)         ' full path to database file
            Dim SplitFileName_DB() As String
            SplitFileName_DB = Split(mainForm.sFilePath, "\")
            name = SplitFileName_DB(SplitFileName_DB.Count - 1)
            SplitFileName_DB = Split(name, ".")
            name = SplitFileName_DB(0)

            mainForm.fileNames.Add(name)             ' add name of each file to name collection


            '   Create collection of Excel files workSheets

            Dim ws As ExcelWorksheet
            Dim excelFile = New FileInfo(mainForm.sFilePath)

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            Dim Excel As ExcelPackage = New ExcelPackage(excelFile)


            key = key + 1

            mainForm.i_wsDict = New Dictionary(Of Integer, ExcelWorksheet)

            For i As Integer = 0 To Excel.Workbook.Worksheets.Count - 1

                mainForm.i_xlTableDict = New Dictionary(Of Integer, ExcelTable)

                ws = Excel.Workbook.Worksheets(i)

                mainForm.i_wsDict.Add(i, ws)

                Dim k As Integer = 0
                For Each tbl As ExcelTable In ws.Tables

                    mainForm.i_xlTableDict.Add(k, tbl)
                    k = k + 1
                Next tbl

                mainForm.i_pivotTableDict.Add(i, mainForm.i_xlTableDict)
            Next i

            mainForm.i_pivot_wsDict.Add(key - 1, mainForm.i_wsDict)
            mainForm.i_superPivotDict.Add(key - 1, mainForm.i_pivotTableDict)

            mainForm.i_pivotTableDict = New Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable))
        Next


    End Sub

    '===================================================================================      
    '                === Create datatable ===
    '===================================================================================

    Sub create_dt()

        Dim dt As DataTable
        Dim xlTable As ExcelTable
        Dim adr As String
        Dim rng As ExcelRange
        Dim r_xlTable, c_xlTable As Integer




    End Sub


    '===================================================================================
    '             === Create dataset ===
    '===================================================================================

    Sub create_dataset()

        Dim dt As DataTable

        Dim xlTable As ExcelTable
        Dim adr As String
        Dim row As DataRow
        Dim ws As ExcelWorksheet
        Dim r_xlTable, c_xlTable As Integer
        Dim rng As ExcelRange


        mainForm.dts = New DataSet

        ws = mainForm.i_pivot_wsDict(mainForm.iDepartment)(mainForm.iCategory)

        For k As Integer = 1 To ws.Tables.Count - 1

            dt = New DataTable

            xlTable = ws.Tables(k)
            c_xlTable = xlTable.Address.Columns
            r_xlTable = xlTable.Address.Rows

            adr = xlTable.Address.Address
            rng = ws.Cells(adr)

            'Adding the Columns
            For i = 0 To c_xlTable - 1
                dt.Columns.Add(rng.Value(0, i))
            Next i

            dt.TableName = xlTable.Name

            dt.Columns(0).DataType = System.Type.GetType("System.Int32")
            dt.Columns(1).DataType = System.Type.GetType("System.String")
            dt.Columns(2).DataType = System.Type.GetType("System.Int32")
            dt.Columns(3).DataType = System.Type.GetType("System.String")
            dt.Columns(4).DataType = System.Type.GetType("System.Int32")
            dt.Columns(5).DataType = System.Type.GetType("System.String")
            dt.Columns(6).DataType = System.Type.GetType("System.Int32")
            dt.Columns(7).DataType = System.Type.GetType("System.String")
            dt.Columns(8).DataType = System.Type.GetType("System.Int32")

            'Add Rows from Excel table

            For i = 1 To r_xlTable - 2
                row = dt.Rows.Add()

                For j = 0 To c_xlTable - 2

                    If rng.Value(i, j) = Nothing Then
                        Select Case j
                            Case 3
                                row.Item(j) = ""
                            Case 4
                                row.Item(j) = 0
                            Case 5
                                row.Item(j) = ""
                            Case 6
                                row.Item(j) = 0
                            Case 7
                                row.Item(j) = ""
                            Case 8
                                row.Item(j) = 0
                        End Select
                    Else
                        row.Item(j) = rng.Value(i, j)
                    End If

                Next j
            Next i
            mainForm.dts.Tables.Add(dt)
        Next k
    End Sub
End Module
