Imports System.Text.RegularExpressions
Public Class exportExcel

    Dim exportList As List(Of Integer)
    Private Sub exportExcel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        exportList = New List(Of Integer)
    End Sub
    Private Sub btn_cancelExport_Click(sender As Object, e As EventArgs) Handles btn_cancelExport.Click
        Me.Close()
    End Sub

    Private Sub btn_export_Click(sender As Object, e As EventArgs) Handles btn_export.Click

        exportList.Clear()

        Select Case chbx_lighting.Checked
            Case True
                exportList.Add(0)
        End Select

        Select Case chbx_screen.Checked
            Case True
                exportList.Add(1)
        End Select

        Select Case chbx_commutation.Checked
            Case True
                exportList.Add(2)
        End Select

        Select Case chbx_truss.Checked
            Case True
                exportList.Add(3)
                mainForm.iDepartment = 3
        End Select

        Select Case chbx_constr.Checked
            Case True
                exportList.Add(4)
        End Select

        Select Case chbx_sound.Checked
            Case True
                exportList.Add(5)
        End Select

        Dim format As String = ("yyy MM dd HH':'mm':'ss")
        Dim myDate As DateTime = DateTime.Now
        Dim timeStampDir As String

        timeStampDir = myDate.ToString(format)

        timeStampDir = Regex.Replace(timeStampDir, "\D", "")            ' timestamp name

        For i As Integer = 0 To exportList.Count - 1
            exportDataset(exportList(i), timeStampDir)
        Next i

        Select Case MsgBox("Файлы Excel находятся в папке " & mainForm.exportDir & vbCrLf & "Открыть ее?" _
               , vbYesNo + vbInformation)
            Case MsgBoxResult.Yes
                Me.Close()
                Process.Start("explorer.exe", mainForm.exportDir)
            Case MsgBoxResult.No
                Me.Close()
        End Select

    End Sub

    Private Sub chbx_all_CheckedChanged(sender As Object, e As EventArgs) Handles chbx_all.CheckedChanged
        Select Case chbx_all.Checked
            Case True
                chbx_lighting.Checked = True
                chbx_screen.Checked = True
                chbx_commutation.Checked = True
                chbx_truss.Checked = True
                chbx_constr.Checked = True
                chbx_sound.Checked = True

            Case False
                chbx_lighting.Checked = False
                chbx_screen.Checked = False
                chbx_commutation.Checked = False
                chbx_truss.Checked = False
                chbx_constr.Checked = False
                chbx_sound.Checked = False
                exportList.Clear()
        End Select
    End Sub


End Class