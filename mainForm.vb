Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO


Public Class mainForm

    Public sDir As String
    Public sFilePath As String

    Dim currentDate As Date = Date.Now
    Dim lastRunDate As Date = My.Settings.lastRun

    Public fileNames As Collection
    Public filePath As Collection

    ' Dictionaries with Integer key
    Public i_wsDict As Dictionary(Of Integer, ExcelWorksheet)
    Public i_xlTableDict As Dictionary(Of Integer, ExcelTable)
    Public i_pivot_wsDict As Dictionary(Of Integer, Dictionary(Of Integer, ExcelWorksheet))
    Public i_pivotTableDict As Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable))
    Public i_superPivotDict As Dictionary(Of Integer, Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable)))

    Public iDepartment, iCategory, iCompany As Integer

    Public dts As DataSet

    Public sCompany() As String = {"belimlight", "PRLighting", "blackout", "vision", "stage"}

    Public delta As Integer     ' to increase or decrease table when push Add or Delete 

    '===================================================================================
    '             === mainForm_Load ===
    '===================================================================================
    Private Sub mainForm_Load(sender As Object, e As EventArgs) Handles Me.Load

        '                   check expiration date
        '-----------------------------------------------------------------------------------

        Dim daysStayed As Int32 = My.Settings.expireDate.Subtract(currentDate).Days

        menuItem_department.Enabled = False
        menuItem_company.Enabled = False

        If lastRunDate.Subtract(currentDate).Days > 0 Then
            MsgBox("Check date and time settings!")
            Me.Close()
        Else
            My.Settings.lastRun = currentDate
            My.Settings.Save()
        End If

        If daysStayed > 0 Then
            Return
        Else
            MsgBox("This app has expired!")
            Me.Close()
        End If
    End Sub
    '===================================================================================
    '             === Menu items ===
    '===================================================================================
#Region "Menu items"

    '                   loadDataBase
    '-----------------------------------------------------------------------------------
    Private Sub FolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FolderToolStripMenuItem.Click

        iDepartment = 0
        iCategory = 0
        iCompany = 1
        loadDataBaseFolder()
        menuItem_department.Enabled = True
        menuItem_company.Enabled = True

    End Sub
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

#Region "select Lighting"
    Private Sub item_movHeads_Click(sender As Object, e As EventArgs) Handles item_movHeads.Click

        iDepartment = 0
        iCategory = 0
        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_strobes_Click(sender As Object, e As EventArgs) Handles item_strobes.Click

        iDepartment = 0
        iCategory = 1
        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_blinders_Click(sender As Object, e As EventArgs) Handles item_blinders.Click

        iDepartment = 0
        iCategory = 2
        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_arch_Click(sender As Object, e As EventArgs) Handles item_arch.Click

        iDepartment = 0
        iCategory = 3
        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_LED_Click(sender As Object, e As EventArgs) Handles item_LED.Click

        iDepartment = 0
        iCategory = 4
        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_smoke_Click(sender As Object, e As EventArgs) Handles item_smoke.Click

        iDepartment = 0
        iCategory = 5
        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_consoles_Click(sender As Object, e As EventArgs) Handles item_consoles.Click

        iDepartment = 0
        iCategory = 6
        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_intercom_Click(sender As Object, e As EventArgs) Handles item_intercom.Click

        iDepartment = 0
        iCategory = 7
        writeToLabel("Lighting", sender)

    End Sub
#End Region

#Region "select Screen"
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

#Region "Select Commutation"
    Private Sub item_pwrdistr_Click(sender As Object, e As EventArgs) Handles item_pwrdistr.Click
        iDepartment = 2
        iCategory = 0
        writeToLabel("Commutation", sender)
    End Sub

    Private Sub item_comm_Click(sender As Object, e As EventArgs) Handles item_comm.Click
        iDepartment = 2
        iCategory = 1
        writeToLabel("Commutation", sender)
    End Sub

    Private Sub item_pwrcomm_Click(sender As Object, e As EventArgs) Handles item_pwrcomm.Click
        iDepartment = 2
        iCategory = 2
        writeToLabel("Commutation", sender)
    End Sub

    Private Sub item_rest_Click(sender As Object, e As EventArgs) Handles item_rest.Click
        iDepartment = 2
        iCategory = 3
        writeToLabel("Commutation", sender)
    End Sub
#End Region

#Region "Select Truss and motors"
    Private Sub item_truss30x30_Click(sender As Object, e As EventArgs) Handles item_truss30x30.Click
        iDepartment = 3
        iCategory = 0
        writeToLabel("Trusses and motors", sender)
    End Sub

    Private Sub item_truss40x40_Click(sender As Object, e As EventArgs) Handles item_truss40x40.Click
        iDepartment = 3
        iCategory = 1
        writeToLabel("Trusses and motors", sender)
    End Sub

    Private Sub item_truss50x60_Click(sender As Object, e As EventArgs) Handles item_truss50x60.Click
        iDepartment = 3
        iCategory = 2
        writeToLabel("Trusses and motors", sender)
    End Sub

    Private Sub item_motors_Click(sender As Object, e As EventArgs) Handles item_motors.Click
        iDepartment = 3
        iCategory = 3
        writeToLabel("Trusses and motors", sender)
    End Sub

    Private Sub item_rigging_Click(sender As Object, e As EventArgs) Handles item_rigging.Click
        iDepartment = 3
        iCategory = 4
        writeToLabel("Trusses and motors", sender)
    End Sub

    Private Sub item_diff_Click(sender As Object, e As EventArgs) Handles item_diff.Click
        iDepartment = 3
        iCategory = 5
        writeToLabel("Trusses and motors", sender)
    End Sub

    Private Sub item_completeConstr_Click(sender As Object, e As EventArgs) Handles item_completeConstr.Click
        iDepartment = 3
        iCategory = 6
        writeToLabel("Trusses and motors", sender)
    End Sub

    Private Sub item_stagelifts_Click(sender As Object, e As EventArgs) Handles item_stagelifts.Click
        iDepartment = 3
        iCategory = 7
        writeToLabel("Trusses and motors", sender)
    End Sub
#End Region

#Region "Select Construction"
    Private Sub item_stageModules_Click(sender As Object, e As EventArgs) Handles item_stageModules.Click
        iDepartment = 4
        iCategory = 0
        writeToLabel("Construction", sender)
    End Sub

    Private Sub item_scaffold_J001_Click(sender As Object, e As EventArgs) Handles item_scaffold_J001.Click
        iDepartment = 4
        iCategory = 1
        writeToLabel("Construction", sender)
    End Sub

    Private Sub item_scaffold_J004_Click(sender As Object, e As EventArgs) Handles item_scaffold_J004.Click
        iDepartment = 4
        iCategory = 2
        writeToLabel("Construction", sender)
    End Sub

    Private Sub item_scaffold_steps_Click(sender As Object, e As EventArgs) Handles item_scaffold_steps.Click
        iDepartment = 4
        iCategory = 3
        writeToLabel("Construction", sender)
    End Sub

    Private Sub item_barricades_Click(sender As Object, e As EventArgs) Handles item_barricades.Click
        iDepartment = 4
        iCategory = 4
        writeToLabel("Construction", sender)
    End Sub

    Private Sub item_details_Click(sender As Object, e As EventArgs) Handles item_details.Click
        iDepartment = 4
        iCategory = 5
        writeToLabel("Construction", sender)
    End Sub
#End Region

#Region "Select Sound"
    Private Sub item_speakers_Click(sender As Object, e As EventArgs) Handles item_speakers.Click
        iDepartment = 5
        iCategory = 0
        writeToLabel("Sound", sender)
    End Sub

    Private Sub item_ampracks_Click(sender As Object, e As EventArgs) Handles item_ampracks.Click
        iDepartment = 5
        iCategory = 1
        writeToLabel("Sound", sender)
    End Sub

    Private Sub item_monitors_Click(sender As Object, e As EventArgs) Handles item_monitors.Click
        iDepartment = 5
        iCategory = 2
        writeToLabel("Sound", sender)
    End Sub

    Private Sub item_mixdesks_Click(sender As Object, e As EventArgs) Handles item_mixdesks.Click
        iDepartment = 5
        iCategory = 3
        writeToLabel("Sound", sender)
    End Sub

    Private Sub item_dj_stage_Click(sender As Object, e As EventArgs) Handles item_dj_stage.Click
        iDepartment = 5
        iCategory = 4
        writeToLabel("Sound", sender)
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
        '   Check is form running
        For Each f As Form In Application.OpenForms
            If f.Name = "sumForm" Then
                sumForm.dgv_sum.DataSource = dts.Tables(0)
                format_sumDGV()
            End If
        Next f

    End Sub

    Private Sub item_PRLighting_Click(sender As Object, e As EventArgs) Handles item_PRLighting.Click
        iCompany = 2
        create_dataset()
        writeToLabelCompany(sender)
        Dim c As Color = Color.FromArgb(221, 235, 247)
        dgv.DataSource = dts.Tables(iCompany)
        format_dgv_dataset(c)
        '   Check is form running
        For Each f As Form In Application.OpenForms
            If f.Name = "sumForm" Then
                sumForm.dgv_sum.DataSource = dts.Tables(0)
                format_sumDGV()
            End If
        Next f

    End Sub

    Private Sub item_blackout_Click(sender As Object, e As EventArgs) Handles item_blackout.Click
        iCompany = 3
        create_dataset()
        writeToLabelCompany(sender)
        Dim c As Color = Color.FromArgb(237, 237, 237)
        dgv.DataSource = dts.Tables(iCompany)
        format_dgv_dataset(c)
        '   Check is form running
        For Each f As Form In Application.OpenForms
            If f.Name = "sumForm" Then
                sumForm.dgv_sum.DataSource = dts.Tables(0)
                format_sumDGV()
            End If
        Next f

    End Sub

    Private Sub item_vision_Click(sender As Object, e As EventArgs) Handles item_vision.Click
        iCompany = 4
        create_dataset()
        writeToLabelCompany(sender)
        Dim c As Color = Color.FromArgb(226, 239, 218)
        dgv.DataSource = dts.Tables(iCompany)
        format_dgv_dataset(c)
        '   Check is form running
        For Each f As Form In Application.OpenForms
            If f.Name = "sumForm" Then
                sumForm.dgv_sum.DataSource = dts.Tables(0)
                format_sumDGV()
            End If
        Next f

    End Sub

    Private Sub item_stage_Click(sender As Object, e As EventArgs) Handles item_stage.Click
        iCompany = 5
        create_dataset()
        writeToLabelCompany(sender)
        Dim c As Color = Color.FromArgb(237, 226, 246)
        dgv.DataSource = dts.Tables(iCompany)
        format_dgv_dataset(c)
        '   Check is form running
        For Each f As Form In Application.OpenForms
            If f.Name = "sumForm" Then
                sumForm.dgv_sum.DataSource = dts.Tables(0)
                format_sumDGV()
            End If
        Next f

    End Sub
    Private Sub item_summary_Click(sender As Object, e As EventArgs) Handles item_summary.Click

        'iCompany = 0
        sumForm.Show()
        'create_dataset()
        sumForm.dgv_sum.DataSource = dts.Tables(0)
        sumForm.dgv_sum.Columns(8).Visible = False
        sumForm.dgv_sum.Columns(9).Visible = False
        sumForm.dgv_sum.Columns(10).Visible = False
        format_sumDGV()

    End Sub

#End Region

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
        sumForm.dgv_sum.DataSource = Nothing
    End Sub



    Sub writeToLabelCompany(_sender As Object)
        Me.GroupBox3.Visible = True
        Me.lbl_companyValue.Text = _sender.text
    End Sub

    Private Sub dgv_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv.CellClick
        dgv_clickCell(sender, e)
        calcQuantity()

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

    Private Sub btn_add_Click(sender As Object, e As EventArgs) Handles btn_add.Click
        delta = 1
        addForm.Show()
    End Sub

    Private Sub btn_update_Click(sender As Object, e As EventArgs) Handles btn_update.Click
        updateData()
        calcQuantity()
        format_sumDGV()
        delta = 0
    End Sub

    Private Sub btn_delete_Click(sender As Object, e As EventArgs) Handles btn_delete.Click
        deleteRow()
        delta = -1
    End Sub

    Private Sub btn_save_Click(sender As Object, e As EventArgs) Handles btn_save.Click, btn_cancel.Click
        saveButton(delta)
    End Sub

    '===================================================================================
    '             === Format sumDGV ===
    '===================================================================================

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
        sumForm.dgv_sum.Columns(11).DefaultCellStyle.Font = New Font("Tahoma", 10, FontStyle.Bold)

        sumForm.dgv_sum.Columns(3).DefaultCellStyle.BackColor = col(0)
        sumForm.dgv_sum.Columns(4).DefaultCellStyle.BackColor = col(1)
        sumForm.dgv_sum.Columns(5).DefaultCellStyle.BackColor = col(2)
        sumForm.dgv_sum.Columns(6).DefaultCellStyle.BackColor = col(3)
        sumForm.dgv_sum.Columns(7).DefaultCellStyle.BackColor = col(4)

        For i = 0 To sumForm.dgv_sum.Rows.Count - 2

            sumForm.dgv_sum.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250)
            If sumForm.dgv_sum.Item(11, i).Value = 0 Then
                sumForm.dgv_sum.Item(0, i).Style.BackColor = Color.FromArgb(216, 238, 192)
                sumForm.dgv_sum.Item(1, i).Style.BackColor = Color.FromArgb(216, 238, 192)
                sumForm.dgv_sum.Item(2, i).Style.BackColor = Color.FromArgb(216, 238, 192)
                sumForm.dgv_sum.Item(11, i).Style.BackColor = Color.FromArgb(216, 238, 192)
            Else
                sumForm.dgv_sum.Item(0, i).Style.BackColor = Color.FromArgb(255, 183, 183)
                sumForm.dgv_sum.Item(1, i).Style.BackColor = Color.FromArgb(255, 183, 183)
                sumForm.dgv_sum.Item(2, i).Style.BackColor = Color.FromArgb(255, 183, 183)
                sumForm.dgv_sum.Item(11, i).Style.BackColor = Color.FromArgb(255, 183, 183)
            End If
        Next i

        sumForm.dgv_sum.Columns(8).Visible = False
        sumForm.dgv_sum.Columns(9).Visible = False
        sumForm.dgv_sum.Columns(10).Visible = False

    End Sub
    '===================================================================================      
    '                === Test button ===
    '===================================================================================
#Region "TestButton"
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
        'sumForm.Show()
        'create_dataset()

        'sumForm.dgv_sum.DataSource = dts.Tables(0)
        '-----------------------------------------------------------------------------------------

        'sumForm.dgv_sum.Columns(0).Visible = False

        '-----------------------------------------------------------------------------------------
        'backUp_db()
        '-----------------------------------------------------------------------------------------

    End Sub
#End Region

End Class