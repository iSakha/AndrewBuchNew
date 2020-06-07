Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO
Imports System.Text.RegularExpressions
Imports Ionic.Zip

Module myFunc


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

        For k As Integer = 0 To ws.Tables.Count - 1

            xlTable = ws.Tables(k)
            c_xlTable = xlTable.Address.Columns
            r_xlTable = xlTable.Address.Rows

            adr = xlTable.Address.Address
            rng = ws.Cells(adr)

            Select Case k

                Case = 0

                    dt = New DataTable
                    dt.TableName = xlTable.Name

                    'Adding the Columns
                    For i = 0 To c_xlTable - 1
                        dt.Columns.Add(rng.Value(0, i))
                    Next i

                    dt.Columns(0).DataType = System.Type.GetType("System.Int32")               ' #
                    dt.Columns(1).DataType = System.Type.GetType("System.String")              ' Fixture
                    dt.Columns(2).DataType = System.Type.GetType("System.Int32")               ' Q-ty
                    dt.Columns(3).DataType = System.Type.GetType("System.Int32")               ' BelImlight
                    dt.Columns(4).DataType = System.Type.GetType("System.Int32")               ' PRLightigTouring
                    dt.Columns(5).DataType = System.Type.GetType("System.Int32")               ' BlackOut
                    dt.Columns(6).DataType = System.Type.GetType("System.Int32")               ' Vision
                    dt.Columns(7).DataType = System.Type.GetType("System.Int32")               ' Stage
                    dt.Columns(8).DataType = System.Type.GetType("System.Int32")               ' Weight
                    If mainForm.iDepartment = 3 Then
                        dt.Columns(9).DataType = System.Type.GetType("System.String")          ' Power/length
                    Else
                        dt.Columns(9).DataType = System.Type.GetType("System.Int32")           ' Power/length
                    End If

                    dt.Columns(10).DataType = System.Type.GetType("System.Int32")              ' Price
                    dt.Columns.Add()
                    dt.Columns(11).DataType = System.Type.GetType("System.Int32")              ' Result
                    dt.Columns(11).ColumnName = "Result"


                    For i = 1 To r_xlTable - 1

                        row = dt.Rows.Add()

                        For j = 0 To c_xlTable - 2

                            row.Item(j) = rng.Value(i, j)

                        Next j

                        Dim val, val_bel, val_pr, val_black, val_vis, val_st As Integer

                        val = row.Item(2)
                        val_bel = row.Item(3)
                        val_pr = row.Item(4)
                        val_black = row.Item(5)
                        val_vis = row.Item(6)
                        val_st = row.Item(7)

                        row.Item(c_xlTable) = val - (val_bel + val_pr + val_black + val_vis + val_st)

                    Next i

                Case > 0

                    dt = New DataTable
                    dt.TableName = xlTable.Name

                    'Adding the Columns
                    For i = 0 To c_xlTable - 1
                        dt.Columns.Add(rng.Value(0, i))
                    Next i

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

                    For i = 1 To r_xlTable - 1
                        row = dt.Rows.Add()

                        For j = 0 To c_xlTable - 1

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

            End Select

            mainForm.dts.Tables.Add(dt)
        Next k
    End Sub

    '===================================================================================
    '             === CellClick on DGV ===
    '===================================================================================
    Sub dgv_clickCell(_sender As Object, _e As DataGridViewCellEventArgs)

        Dim index As Integer
        index = _e.RowIndex
        Console.WriteLine(_e)
        Dim selectedRow As DataGridViewRow
        selectedRow = _sender.Rows(index)

        mainForm.rtb_fixtureName.Text = selectedRow.Cells(1).Value.ToString
        mainForm.txt_qty.Text = selectedRow.Cells(2).Value.ToString
        mainForm.rtb_FirstName.Text = selectedRow.Cells(3).Value.ToString
        mainForm.txt_qty1.Text = selectedRow.Cells(4).Value.ToString
        mainForm.rtb_SecondName.Text = selectedRow.Cells(5).Value.ToString
        mainForm.txt_qty2.Text = selectedRow.Cells(6).Value.ToString
        mainForm.rtb_ThirdName.Text = selectedRow.Cells(7).Value.ToString
        mainForm.txt_qty3.Text = selectedRow.Cells(8).Value.ToString

        mainForm.dgv.Rows(index).Selected = True
        '   Check is form running
        For Each f As Form In Application.OpenForms
            If f.Name = "sumForm" Then
                sumForm.dgv_sum.ClearSelection()
                sumForm.dgv_sum.Rows(index).Selected = True
            End If
        Next f
    End Sub
#Region "next/prev buttons"
    '===================================================================================
    '             === Prev record ===
    '===================================================================================
    Sub prevRecord()

        Dim index As Integer
        Dim selectedRow As DataGridViewRow


        index = mainForm.dgv.CurrentRow.Index

        mainForm.dgv.ClearSelection()

        mainForm.dgv.CurrentCell = mainForm.dgv.Item(0, index)
        mainForm.dgv.Rows(index).Selected = True


        If index = 0 Then
            index = mainForm.dgv.Rows.Count - 1
        End If



        index = index - 1
        mainForm.dgv.CurrentCell = mainForm.dgv.Item(0, index)
        mainForm.dgv.Rows(index).Selected = True
        '   Check is form running
        For Each f As Form In Application.OpenForms
            If f.Name = "sumForm" Then
                sumForm.dgv_sum.ClearSelection()
                sumForm.dgv_sum.Rows(index).Selected = True
            End If
        Next f

        selectedRow = mainForm.dgv.Rows(index)

        mainForm.rtb_fixtureName.Text = selectedRow.Cells(1).Value.ToString
        mainForm.txt_qty.Text = selectedRow.Cells(2).Value.ToString
        mainForm.rtb_FirstName.Text = selectedRow.Cells(3).Value.ToString
        mainForm.txt_qty1.Text = selectedRow.Cells(4).Value.ToString
        mainForm.rtb_SecondName.Text = selectedRow.Cells(5).Value.ToString
        mainForm.txt_qty2.Text = selectedRow.Cells(6).Value.ToString
        mainForm.rtb_ThirdName.Text = selectedRow.Cells(7).Value.ToString
        mainForm.txt_qty3.Text = selectedRow.Cells(8).Value.ToString


        'calcQuantity()

    End Sub

    '===================================================================================
    '             === Next record ===
    '===================================================================================
    Sub nextRecord()
        Dim index As Integer
        Dim selectedRow As DataGridViewRow

        index = mainForm.dgv.CurrentRow.Index

        mainForm.dgv.ClearSelection()


        mainForm.dgv.CurrentCell = mainForm.dgv.Item(0, index)
        mainForm.dgv.Rows(index).Selected = True

        If index = mainForm.dgv.Rows.Count - 2 Then
            index = -1
        End If



        index = index + 1
        mainForm.dgv.CurrentCell = mainForm.dgv.Item(0, index)
        mainForm.dgv.Rows(index).Selected = True
        '   Check is form running
        For Each f As Form In Application.OpenForms
            If f.Name = "sumForm" Then
                sumForm.dgv_sum.ClearSelection()
                sumForm.dgv_sum.Rows(index).Selected = True
            End If
        Next f

        selectedRow = mainForm.dgv.Rows(index)

        mainForm.rtb_fixtureName.Text = selectedRow.Cells(1).Value.ToString
        mainForm.txt_qty.Text = selectedRow.Cells(2).Value.ToString
        mainForm.rtb_FirstName.Text = selectedRow.Cells(3).Value.ToString
        mainForm.txt_qty1.Text = selectedRow.Cells(4).Value.ToString
        mainForm.rtb_SecondName.Text = selectedRow.Cells(5).Value.ToString
        mainForm.txt_qty2.Text = selectedRow.Cells(6).Value.ToString
        mainForm.rtb_ThirdName.Text = selectedRow.Cells(7).Value.ToString
        mainForm.txt_qty3.Text = selectedRow.Cells(8).Value.ToString

    End Sub
#End Region

    '===================================================================================
    '             === Calculate quantity ===
    '===================================================================================
    Sub calcQuantity()

        Dim index As Integer
        Dim i, j, qty, sum As Integer
        ' Dim smetaQty, companiesQty As Integer

        'i = mainForm.iCategory + 1

        index = mainForm.dgv.CurrentRow.Index

        For j = 1 To mainForm.dts.Tables.Count - 1
            sum = 0
            qty = mainForm.dts.Tables(j).Rows(index).Item(4)
            sum = sum + qty
            qty = mainForm.dts.Tables(j).Rows(index).Item(6)
            sum = sum + qty
            qty = mainForm.dts.Tables(j).Rows(index).Item(8)
            sum = sum + qty

            mainForm.dgv_result.Rows(0).Cells(j).Value = sum


        Next j

        Dim smetaQty As Integer = mainForm.dts.Tables(0).Rows(index).Item(2)

        Dim companiesQty As Integer = mainForm.dgv_result.Rows(0).Cells(1).Value +
        mainForm.dgv_result.Rows(0).Cells(2).Value +
        mainForm.dgv_result.Rows(0).Cells(3).Value +
        mainForm.dgv_result.Rows(0).Cells(4).Value +
        mainForm.dgv_result.Rows(0).Cells(5).Value

        mainForm.dgv_result.Rows(0).Cells(0).Value = smetaQty
        mainForm.dgv_result.Rows(0).Cells(6).Value = smetaQty - companiesQty

        If (smetaQty - companiesQty = 0) Then
            mainForm.dgv_result.Item(6, 0).Style.BackColor = Color.LightGreen
        Else
            mainForm.dgv_result.Item(6, 0).Style.BackColor = Color.LightPink
        End If
    End Sub

    '===================================================================================
    '             === UPDATE data in DB ===
    '===================================================================================
    Sub updateData()

        Dim row As DataRow
        Dim dt As DataTable
        dt = mainForm.dts.Tables(mainForm.iCompany)
        Dim index As Integer = mainForm.dgv.CurrentRow.Index
        row = dt.Rows(index)


        Dim sRow() As String

        sRow = New String() {
                mainForm.rtb_fixtureName.Text,
                mainForm.txt_qty.Text,
                mainForm.rtb_FirstName.Text,
                mainForm.txt_qty1.Text,
                mainForm.rtb_SecondName.Text,
                mainForm.txt_qty2.Text,
                mainForm.rtb_ThirdName.Text,
                mainForm.txt_qty3.Text
            }

        '   Chek null values in textboxes

        For i As Integer = 1 To sRow.Count - 1 Step 2
            If sRow(i) = "" Then
                MsgBox("Поле количества приборов не может быть пустым!")
                mainForm.btn_save.Enabled = False
                Exit Sub
            End If
        Next i
        Console.WriteLine(dt.TableName)
        For i As Integer = 1 To mainForm.dts.Tables.Count

            For j As Integer = 1 To mainForm.dts.Tables(1).Columns.Count - 1
                row.Item(j) = sRow(j - 1)
            Next j

        Next i

        Dim qty As Integer
        Dim qty_belimlight, qty_pr, qty_black, qty_vis, qty_stage As Integer
        qty = dt.Rows(index).Item(4) + dt.Rows(index).Item(6) + dt.Rows(index).Item(8)
        Console.WriteLine(qty)

        mainForm.dts.Tables(0).Rows(index).Item(mainForm.iCompany + 2) = qty

        qty_belimlight = mainForm.dts.Tables(0).Rows(index).Item(3)
        qty_pr = mainForm.dts.Tables(0).Rows(index).Item(4)
        qty_black = mainForm.dts.Tables(0).Rows(index).Item(5)
        qty_vis = mainForm.dts.Tables(0).Rows(index).Item(6)
        qty_stage = mainForm.dts.Tables(0).Rows(index).Item(7)

        mainForm.dts.Tables(0).Rows(index).Item(11) = mainForm.dts.Tables(0).Rows(index).Item(2) - (qty_belimlight + qty_pr + qty_black + qty_vis + qty_stage)
        mainForm.dts.AcceptChanges()
        mainForm.dgv.DataSource = mainForm.dts.Tables(mainForm.iCompany)
        sumForm.dgv_sum.DataSource = mainForm.dts.Tables(0)


    End Sub
    '===================================================================================
    '             === DELETE data from DB ===
    '===================================================================================
    Sub deleteRow()

        Dim index As Integer = mainForm.dgv.CurrentRow.Index
        Dim row As DataRow

        For Each dt As DataTable In mainForm.dts.Tables
            row = dt.Rows(index)
            row.Delete()
        Next dt
    End Sub
    '===================================================================================
    '             === SAVE data to DB ===
    '===================================================================================

    Sub saveButton(_delta As Integer)

        Dim oldAddr As OfficeOpenXml.ExcelAddressBase
        Dim newAddr As OfficeOpenXml.ExcelAddressBase

        Dim dt As DataTable
        Dim ws As ExcelWorksheet

        Dim xlTable As ExcelTable
        Dim startCellAddress As String

        Dim excelFile = New FileInfo(mainForm.filePath(mainForm.iDepartment + 1))
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Dim Excel As ExcelPackage = New ExcelPackage(excelFile)

        ws = Excel.Workbook.Worksheets(mainForm.iCategory)

        Select Case _delta
            Case 0
                '   Write to Exceltable
                dt = mainForm.dts.Tables(mainForm.iCompany)

                xlTable = ws.Tables(mainForm.iCompany)
                startCellAddress = xlTable.Range.Start.Address

                xlTable.Range.Clear()
                ws.Cells(startCellAddress).LoadFromDataTable(dt, True)

                '   Write to pivot Exceltable
                dt = mainForm.dts.Tables(0)

                xlTable = ws.Tables(0)
                startCellAddress = xlTable.Range.Start.Address

                xlTable.Range.Clear()
                ws.Cells(startCellAddress).LoadFromDataTable(dt, True)

            Case <> 0
                For i As Integer = 1 To mainForm.sCompany.Count

                    dt = mainForm.dts.Tables(i)

                    xlTable = ws.Tables(i)
                    oldAddr = xlTable.Address
                    xlTable.Range.Clear()
                    newAddr = New ExcelAddressBase(oldAddr.Start.Row, oldAddr.Start.Column, oldAddr.End.Row + _delta, oldAddr.End.Column)
                    xlTable.TableXml.InnerXml = xlTable.TableXml.InnerXml.Replace(oldAddr.ToString(), newAddr.ToString())

                    startCellAddress = xlTable.Range.Start.Address

                    ws.Cells(startCellAddress).LoadFromDataTable(dt, True)

                Next i

                '   Write to pivot Exceltable
                dt = mainForm.dts.Tables(0)

                xlTable = ws.Tables(0)
                oldAddr = xlTable.Address
                newAddr = New ExcelAddressBase(oldAddr.Start.Row, oldAddr.Start.Column, oldAddr.End.Row + _delta, oldAddr.End.Column)
                xlTable.TableXml.InnerXml = xlTable.TableXml.InnerXml.Replace(oldAddr.ToString(), newAddr.ToString())

                startCellAddress = xlTable.Range.Start.Address

                xlTable.Range.Clear()
                ws.Cells(startCellAddress).LoadFromDataTable(dt, True)

        End Select

        'formatXl_table()

        Excel.SaveAs(excelFile)

        compressFiles()

    End Sub
    '===================================================================================
    '             === block CDUDbuttons ===
    '===================================================================================
    Sub blockButtons()

        mainForm.btn_add.Enabled = False
        mainForm.btn_update.Enabled = False
        mainForm.btn_delete.Enabled = False
        mainForm.btn_next.Enabled = False
        mainForm.btn_prev.Enabled = False

        mainForm.menuItem_department.Enabled = False
        mainForm.menuItem_company.Enabled = False

        mainForm.btn_save.FlatStyle = FlatStyle.Flat
        mainForm.btn_cancel.FlatStyle = FlatStyle.Flat
    End Sub
    '===================================================================================
    '             === unblock CDUDbuttons ===
    '===================================================================================
    Sub unBlockButtons()

        mainForm.btn_add.Enabled = True
        mainForm.btn_update.Enabled = True
        mainForm.btn_delete.Enabled = True
        mainForm.btn_next.Enabled = True
        mainForm.btn_prev.Enabled = True

        mainForm.menuItem_department.Enabled = True
        mainForm.menuItem_company.Enabled = True

        mainForm.btn_save.FlatStyle = FlatStyle.Standard
        mainForm.btn_cancel.FlatStyle = FlatStyle.Standard

    End Sub

    '===================================================================================
    '             === Format Excel table ===
    '===================================================================================

    Sub formatXl_table(_sPath As String, _worksheetNumber As Integer)

        Dim excelFile = New FileInfo(_sPath)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Dim Excel As ExcelPackage = New ExcelPackage(excelFile)

        Dim rngHeader, rngSide, rngTbl_0(5) As ExcelRange
        Dim startRow, startColumn, endRow As Integer
        Dim sideBackColor As Color = Color.FromArgb(242, 245, 245)
        Dim ws As ExcelWorksheet
        ws = Excel.Workbook.Worksheets(_worksheetNumber)

        Dim col() As Color

        col = {Color.FromArgb(252, 228, 214), Color.FromArgb(221, 235, 247), Color.FromArgb(237, 237, 237),
            Color.FromArgb(226, 239, 218), Color.FromArgb(237, 226, 246)}

        Dim i As Integer = 0

        For Each tbl As ExcelTable In ws.Tables

            startRow = tbl.Address.Start.Row
            endRow = tbl.Address.End.Row
            startColumn = tbl.Address.Start.Column

            rngHeader = ws.Cells(startRow, startColumn + 3, startRow, startColumn + 8)

            rngSide = ws.Cells(startRow + 1, startColumn + 1, endRow, startColumn + 2)


            rngTbl_0(0) = ws.Cells(startRow, startColumn + 3, endRow, startColumn + 3)
            rngTbl_0(1) = ws.Cells(startRow, startColumn + 4, endRow, startColumn + 4)
            rngTbl_0(2) = ws.Cells(startRow, startColumn + 5, endRow, startColumn + 5)
            rngTbl_0(3) = ws.Cells(startRow, startColumn + 6, endRow, startColumn + 6)
            rngTbl_0(4) = ws.Cells(startRow, startColumn + 7, endRow, startColumn + 7)

            rngSide.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            rngSide.Style.Fill.BackgroundColor.SetColor(sideBackColor)

            rngSide.Style.Font.Size = 11
            rngSide.Style.Font.Italic = True
            rngSide.Style.Font.Bold = True
            rngSide.Style.Font.Name = "Calibri"

            Select Case i
                Case 0

                    For j As Integer = 0 To 4
                        rngTbl_0(j).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                        rngTbl_0(j).Style.Fill.BackgroundColor.SetColor(col(j))
                    Next j

                Case <> 0

                    rngHeader.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    rngHeader.Style.Fill.BackgroundColor.SetColor(col(i - 1))

            End Select

            i = i + 1

        Next tbl

        Excel.SaveAs(excelFile)

    End Sub
    '===================================================================================
    '             === Export dataset ===
    '===================================================================================
    Sub exportDataset()

        Dim columnWidth(11) As Integer
        columnWidth = {4, 52, 9, 42, 25, 37, 11, 44, 11, 17, 13}
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Dim Excel As ExcelPackage = New ExcelPackage()

        Dim ws As ExcelWorksheet
        Dim xlTable As ExcelTable
        Dim xlTableName As String
        Dim startRow(6) As Integer
        Dim rng As ExcelRange
        Dim startColumn, endRow, endColumn, tiltShift As Integer

        tiltShift = 11
        startColumn = 3
        startRow(6) = New Integer()

        startRow(0) = 3
        endRow = startRow(0) + mainForm.dts.Tables(0).Rows.Count
        startRow(1) = endRow + tiltShift
        endRow = startRow(1) + mainForm.dts.Tables(1).Rows.Count
        startRow(2) = endRow + tiltShift
        endRow = startRow(2) + mainForm.dts.Tables(2).Rows.Count
        startRow(3) = endRow + tiltShift
        endRow = startRow(3) + mainForm.dts.Tables(2).Rows.Count
        startRow(4) = endRow + tiltShift
        endRow = startRow(4) + mainForm.dts.Tables(2).Rows.Count
        startRow(5) = endRow + tiltShift

        ws = Excel.Workbook.Worksheets.Add("test")

        For k As Integer = 0 To 10
            ws.Column(k + 3).Width = columnWidth(k)
        Next k

        For i As Integer = 0 To mainForm.dts.Tables.Count - 1

            endRow = startRow(i) + mainForm.dts.Tables(i).Rows.Count
            endColumn = startColumn + mainForm.dts.Tables(i).Columns.Count - 1
            xlTableName = mainForm.dts.Tables(i).TableName
            rng = ws.Cells(startRow(i), startColumn, endRow, endColumn)

            xlTable = ws.Tables.Add(rng, xlTableName)
            ws.Cells("C" & startRow(i)).LoadFromDataTable(mainForm.dts.Tables(i), True)
            xlTable.TableStyle = TableStyles.Light15

        Next i

        Dim exportDir As String = Directory.GetCurrentDirectory() & "\ExcelExport"
        Dim sPath As String = exportDir & "\ExcelExport_DB.xlsx"

        Excel.SaveAs(New FileInfo(sPath))

        formatXl_table(sPath, 0)

    End Sub

End Module
