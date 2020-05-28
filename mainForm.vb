Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO


Public Class mainForm

    Public sDir As String
    Public sFilePath As String

    Public fileNames As Collection

    ' Dictionaries with Integer key
    Public i_wsDict As Dictionary(Of Integer, ExcelWorksheet)
    Public i_xlTableDict As Dictionary(Of Integer, ExcelTable)
    Public i_pivot_wsDict As Dictionary(Of Integer, Dictionary(Of Integer, ExcelWorksheet))
    Public i_pivotTableDict As Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable))
    Public i_superPivotDict As Dictionary(Of Integer, Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable)))

    Public iDepartment, iCategory, iCompany As Integer

    Public dts As DataSet

    Public sCompany() As String = {"belimlight", "PRLighting", "blackout", "vision", "stage"}
    Private Sub FolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FolderToolStripMenuItem.Click

        iDepartment = 0
        iCategory = 0
        iCompany = 1
        loadDataBaseFolder()

    End Sub


    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        'mainForm.btn_loadDB.PerformClick()
        'createLightingDataset()
    End Sub
#Region "select Lighting categories"
    Private Sub item_movHeads_Click(sender As Object, e As EventArgs) Handles item_movHeads.Click

        iDepartment = 0
        iCategory = 0
        'If mainForm.lightDataset Is Nothing Then
        '    createLightingDataset()
        'End If

        writeToLabel("Lighting", sender)



    End Sub

    Private Sub item_strobes_Click(sender As Object, e As EventArgs) Handles item_strobes.Click

        iDepartment = 0
        iCategory = 1
        'If mainForm.lightDataset Is Nothing Then
        '    createLightingDataset()
        'End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_blinders_Click(sender As Object, e As EventArgs) Handles item_blinders.Click

        iDepartment = 0
        iCategory = 2
        'If mainForm.lightDataset Is Nothing Then
        '    createLightingDataset()
        'End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_arch_Click(sender As Object, e As EventArgs) Handles item_arch.Click

        iDepartment = 0
        iCategory = 3
        'If mainForm.lightDataset Is Nothing Then
        '    createLightingDataset()
        'End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_LED_Click(sender As Object, e As EventArgs) Handles item_LED.Click

        iDepartment = 0
        iCategory = 4
        'If mainForm.lightDataset Is Nothing Then
        '    createLightingDataset()
        'End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_smoke_Click(sender As Object, e As EventArgs) Handles item_smoke.Click

        iDepartment = 0
        iCategory = 5
        'If mainForm.lightDataset Is Nothing Then
        '    createLightingDataset()
        'End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_consoles_Click(sender As Object, e As EventArgs) Handles item_consoles.Click

        iDepartment = 0
        iCategory = 6
        'If mainForm.lightDataset Is Nothing Then
        '    createLightingDataset()
        'End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_intercom_Click(sender As Object, e As EventArgs) Handles item_intercom.Click

        iDepartment = 0
        iCategory = 7
        'If mainForm.lightDataset Is Nothing Then
        '    createLightingDataset()
        'End If

        writeToLabel("Lighting", sender)

    End Sub
#End Region
#Region "select Screen categories"
    Private Sub item_modules_Click(sender As Object, e As EventArgs) Handles item_modules.Click
        iDepartment = 1
        iCategory = 0
        writeToLabel("Screen", sender)
    End Sub

    Private Sub item_servers_Click(sender As Object, e As EventArgs) Handles item_servers.Click
        iDepartment = 1
        iCategory = 1
        writeToLabel("Screen", sender)
    End Sub

    Private Sub item_controllers1_Click(sender As Object, e As EventArgs) Handles item_controllers1.Click
        iDepartment = 1
        iCategory = 2
        writeToLabel("Screen", sender)
    End Sub

    Private Sub item_controllers2_Click(sender As Object, e As EventArgs) Handles item_controllers2.Click
        iDepartment = 1
        iCategory = 3
        writeToLabel("Screen", sender)
    End Sub

    Private Sub item_projectors_Click(sender As Object, e As EventArgs) Handles item_projectors.Click
        iDepartment = 1
        iCategory = 4
        writeToLabel("Screen", sender)
    End Sub

    Private Sub item_scr_construction_Click(sender As Object, e As EventArgs) Handles item_scr_construction.Click
        iDepartment = 1
        iCategory = 5
        writeToLabel("Screen", sender)
    End Sub

    Private Sub item_lightDesks_Click(sender As Object, e As EventArgs) Handles item_lightDesks.Click
        iDepartment = 1
        iCategory = 6
        writeToLabel("Screen", sender)
    End Sub

    Private Sub item_cameras_Click(sender As Object, e As EventArgs) Handles item_cameras.Click
        iDepartment = 1
        iCategory = 7
        writeToLabel("Screen", sender)
    End Sub
#End Region
#Region "select Company"
    Private Sub item_belimlight_Click(sender As Object, e As EventArgs) Handles item_belimlight.Click

        iCompany = 1
        create_dataset()
        writeToLabelCompany(sender)
        Dim c As Color = Color.FromArgb(252, 228, 214)
        dgv.DataSource = dts.Tables(iCompany)
        format_dgv_dataset(c)

    End Sub

    Private Sub item_PRLighting_Click(sender As Object, e As EventArgs) Handles item_PRLighting.Click
        iCompany = 2
        create_dataset()
        writeToLabelCompany(sender)
        Dim c As Color = Color.FromArgb(221, 235, 247)
        dgv.DataSource = dts.Tables(iCompany)
        format_dgv_dataset(c)
    End Sub

    Private Sub item_blackout_Click(sender As Object, e As EventArgs) Handles item_blackout.Click
        iCompany = 3
        create_dataset()
        writeToLabelCompany(sender)
        Dim c As Color = Color.FromArgb(237, 237, 237)
        dgv.DataSource = dts.Tables(iCompany)
        format_dgv_dataset(c)
    End Sub

    Private Sub item_vision_Click(sender As Object, e As EventArgs) Handles item_vision.Click
        iCompany = 4
        create_dataset()
        writeToLabelCompany(sender)
        Dim c As Color = Color.FromArgb(226, 239, 218)
        dgv.DataSource = dts.Tables(iCompany)
        format_dgv_dataset(c)
    End Sub

    Private Sub item_stage_Click(sender As Object, e As EventArgs) Handles item_stage.Click
        iCompany = 5
        create_dataset()
        writeToLabelCompany(sender)
        Dim c As Color = Color.FromArgb(237, 226, 246)
        dgv.DataSource = dts.Tables(iCompany)
        format_dgv_dataset(c)
    End Sub
#End Region

    Private Sub btn_prev_Click(sender As Object, e As EventArgs) Handles btn_prev.Click
        prevRecord()
        calcQuantity()
    End Sub

    Private Sub btn_next_Click(sender As Object, e As EventArgs) Handles btn_next.Click
        nextRecord()
        calcQuantity()
    End Sub

    '===================================================================================
    '             === MY FUNCTIONS ===
    '===================================================================================

    Sub writeToLabel(_department As String, _sender As Object)
        Me.GroupBox1.Visible = True
        Me.GroupBox2.Visible = True
        Me.lbl_dpartmentValue.Text = _department
        Me.lbl_subsectionValue.Text = _sender.text
        Me.dgv.DataSource = Nothing
    End Sub



    Sub writeToLabelCompany(_sender As Object)
        Me.GroupBox3.Visible = True
        Me.lbl_companyValue.Text = _sender.text
    End Sub

    Private Sub dgv_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv.CellClick
        dgv_clickCell(sender, e)
        calcQuantity()

    End Sub

    Private Sub item_summary_Click(sender As Object, e As EventArgs) Handles item_summary.Click
        iCompany = 0
        sumForm.Show()
        create_dataset()
        sumForm.dgv_sum.DataSource = dts.Tables(0)
        sumForm.dgv_sum.Columns(8).Visible = False
        sumForm.dgv_sum.Columns(9).Visible = False
        sumForm.dgv_sum.Columns(10).Visible = False
        format_sumDGV()
    End Sub

    '===================================================================================      
    '                === Format DataGridView ===
    '===================================================================================
    Sub format_dgv_dataset(_color As Color)

        dgv.Columns(0).Width = 40                ' #
        dgv.Columns(1).Width = 175               ' Fixture
        dgv.Columns(2).Width = 40                ' Q-ty
        dgv.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns(3).Width = 220               ' BelImlight_1  (PRLightigTouring, BlackOut, Vision, Stage)
        dgv.Columns(4).Width = 40                ' Q-ty_1
        dgv.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns(5).Width = 220               ' BelImlight_2  (PRLightigTouring, BlackOut, Vision, Stage)
        dgv.Columns(6).Width = 40                ' Q-ty_2
        dgv.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns(7).Width = 180               ' BelImlight_3  (PRLightigTouring, BlackOut, Vision, Stage)
        dgv.Columns(8).Width = 40                ' Q-ty_3
        dgv.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        For i = 0 To dgv.Rows.Count - 2

            'mainForm.DGV_in.Rows(i).Cells(1).Value = Date.FromOADate(mainForm.DGV_in.Rows(i).Cells(1).Value)
            dgv.RowsDefaultCellStyle.BackColor = _color
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250)

        Next i

    End Sub
    Sub format_sumDGV()

        Dim col() As Color

        col = {Color.FromArgb(252, 228, 214), Color.FromArgb(221, 235, 247), Color.FromArgb(237, 237, 237),
            Color.FromArgb(226, 239, 218), Color.FromArgb(237, 226, 246)}

        sumForm.dgv_sum.Columns(0).Width = 55                ' #
        sumForm.dgv_sum.Columns(1).Width = 230               ' Fixture
        sumForm.dgv_sum.Columns(2).Width = 65                ' Q-ty
        sumForm.dgv_sum.Columns(2).DefaultCellStyle.Font = New Font("Tahoma", 10)
        sumForm.dgv_sum.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sumForm.dgv_sum.Columns(3).Width = 62                ' BelImlight
        sumForm.dgv_sum.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sumForm.dgv_sum.Columns(4).Width = 62                ' PRLightigTouring
        sumForm.dgv_sum.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sumForm.dgv_sum.Columns(5).Width = 62                ' BlackOut
        sumForm.dgv_sum.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sumForm.dgv_sum.Columns(6).Width = 62                ' Vision
        sumForm.dgv_sum.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sumForm.dgv_sum.Columns(7).Width = 62                ' Stage
        sumForm.dgv_sum.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        sumForm.dgv_sum.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sumForm.dgv_sum.Columns(11).Width = 65
        sumForm.dgv_sum.Columns(11).DefaultCellStyle.Font = New Font("Tahoma", 10)

        sumForm.dgv_sum.Columns(3).DefaultCellStyle.BackColor = col(0)
        sumForm.dgv_sum.Columns(4).DefaultCellStyle.BackColor = col(1)
        sumForm.dgv_sum.Columns(5).DefaultCellStyle.BackColor = col(2)
        sumForm.dgv_sum.Columns(6).DefaultCellStyle.BackColor = col(3)
        sumForm.dgv_sum.Columns(7).DefaultCellStyle.BackColor = col(4)

        For i = 0 To sumForm.dgv_sum.Rows.Count - 2

            sumForm.dgv_sum.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250)
            If sumForm.dgv_sum.Item(11, i).Value = 0 Then
                sumForm.dgv_sum.Item(0, i).Style.BackColor = Color.LightGreen
                sumForm.dgv_sum.Item(1, i).Style.BackColor = Color.LightGreen
                sumForm.dgv_sum.Item(2, i).Style.BackColor = Color.LightGreen
                sumForm.dgv_sum.Item(11, i).Style.BackColor = Color.LightGreen
            Else
                sumForm.dgv_sum.Item(0, i).Style.BackColor = Color.LightPink
                sumForm.dgv_sum.Item(1, i).Style.BackColor = Color.LightPink
                sumForm.dgv_sum.Item(2, i).Style.BackColor = Color.LightPink
                sumForm.dgv_sum.Item(11, i).Style.BackColor = Color.LightPink
            End If
        Next i

    End Sub
    '===================================================================================      
    '                === Test button ===
    '===================================================================================
    Private Sub btn_test_Click(sender As Object, e As EventArgs) Handles btn_test.Click

        '-----------------------------------------------------------------------------------------
        'Dim xlTable As ExcelTable

        'xlTable = i_superPivotDict(iDepartment)(iCategory)(iCompany)

        'create_dataset()

        'Console.WriteLine(dts.Tables.Count)
        'dgv.DataSource = dts.Tables(0)
        '-----------------------------------------------------------------------------------------

        'dgv_result.Rows.Add()
        'dgv_result.Rows(0).Cells(0).Value = "Hello,world!"

        '-----------------------------------------------------------------------------------------
        'calcQuantity()
        '-----------------------------------------------------------------------------------------
        'dgv_result.Item(6, 0).Style.BackColor = Color.Red
        '-----------------------------------------------------------------------------------------
        sumForm.Show()
        create_dataset()

        sumForm.dgv_sum.DataSource = dts.Tables(0)

    End Sub

End Class