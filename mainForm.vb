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
    Private Sub FolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FolderToolStripMenuItem.Click

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
        writeToLabelCompany(sender)
        Dim c As Color = Color.FromArgb(252, 228, 214)
        'dgv.DataSource = mainForm.lightDataset.Tables(0)
        'format_dgv_dataset(mainForm.tbl_Lighting_tables(0, 0).Name, c)

    End Sub

    Private Sub item_PRLighting_Click(sender As Object, e As EventArgs) Handles item_PRLighting.Click
        iCompany = 2
        writeToLabelCompany(sender)
    End Sub

    Private Sub item_blackout_Click(sender As Object, e As EventArgs) Handles item_blackout.Click
        iCompany = 3
        writeToLabelCompany(sender)
    End Sub

    Private Sub item_vision_Click(sender As Object, e As EventArgs) Handles item_vision.Click
        iCompany = 4
        writeToLabelCompany(sender)
    End Sub

    Private Sub item_stage_Click(sender As Object, e As EventArgs) Handles item_stage.Click
        iCompany = 5
        writeToLabelCompany(sender)
    End Sub
#End Region

    '===================================================================================
    '             === MY FUNCTIONS ===
    '===================================================================================

    Sub writeToLabel(_department As String, _sender As Object)
        Me.GroupBox1.Visible = True
        Me.GroupBox2.Visible = True
        Me.lbl_dpartmentValue.Text = _department
        Me.lbl_subsectionValue.Text = _sender.text
    End Sub



    Sub writeToLabelCompany(_sender As Object)
        Me.GroupBox3.Visible = True
        Me.lbl_companyValue.Text = _sender.text
    End Sub



    '===================================================================================      
    '                === Format DataGridView ===
    '===================================================================================
    Sub format_dgv_dataset(_dtName As String, _color As Color)

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
    '===================================================================================      
    '                === Test button ===
    '===================================================================================
    Private Sub btn_test_Click(sender As Object, e As EventArgs) Handles btn_test.Click

        Dim xlTable As ExcelTable

        xlTable = i_superPivotDict(iDepartment)(iCategory)(iCompany)

        create_dataset()

        Console.WriteLine(dts.Tables.Count)
    End Sub

End Class