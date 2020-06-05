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


    'Sub writeZeroInQtyTxt()
    '    If mainForm.txt_qty.Text = "" Then
    '        mainForm.txt_qty.Text = 0
    '    End If
    '    If mainForm.txt_qty1.Text = "" Then
    '        mainForm.txt_qty1.Text = 0
    '    End If
    '    If mainForm.txt_qty2.Text = "" Then
    '        mainForm.txt_qty2.Text = 0
    '    End If
    '    If mainForm.txt_qty3.Text = "" Then
    '        mainForm.txt_qty3.Text = 0
    '    End If
    'End Sub

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
                    newAddr = New ExcelAddressBase(oldAddr.Start.Row, oldAddr.Start.Column, oldAddr.End.Row + _delta, oldAddr.End.Column)
                    xlTable.TableXml.InnerXml = xlTable.TableXml.InnerXml.Replace(oldAddr.ToString(), newAddr.ToString())

                    startCellAddress = xlTable.Range.Start.Address

                    xlTable.Range.Clear()
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

        Excel.SaveAs(excelFile)

        compressFiles()

    End Sub

    'Sub backUp_db()

    '    Dim folderName, backUpFolder, backUpFile, foundFile As String
    '    Dim format As String = ("yyy MM dd HH':'mm':'ss")
    '    Dim myDate As DateTime = DateTime.Now
    '    folderName = myDate.ToString(format)

    '    Console.WriteLine(folderName)
    '    folderName = Regex.Replace(folderName, "\D", "")            ' timestamp name
    '    Console.WriteLine(folderName)

    '    Console.WriteLine(Directory.GetCurrentDirectory())
    '    backUpFolder = Directory.GetCurrentDirectory() & "\BackUp"
    '    '   Create folder with timestamp name inside backUp folder
    '    My.Computer.FileSystem.CreateDirectory(backUpFolder & "\" & folderName)
    '    backUpFile = Directory.GetCurrentDirectory() & "\BackUp\" & folderName & "\DB.ombckp"

    '    mainForm.FBD.SelectedPath = Directory.GetCurrentDirectory()
    '    If (mainForm.FBD.ShowDialog() = DialogResult.OK) Then
    '        mainForm.sDir = mainForm.FBD.SelectedPath
    '    Else
    '        mainForm.Close()
    '    End If

    '    'For Each foundFile In My.Computer.FileSystem.GetFiles _
    '    '(mainForm.sDir, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.omdb")
    '    '    Console.WriteLine(foundFile)
    '    'Next
    '    'MsgBox("Создаем резервную копию базы данных в папке BackUp", vbOKOnly + vbInformation)
    '    'My.Computer.FileSystem.CopyFile(foundFile, backUpFile)
    'End Sub




End Module
