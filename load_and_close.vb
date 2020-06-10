Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO
Imports System.Text.RegularExpressions
Imports Ionic.Zip

Module load_and_close

    '===================================================================================
    '             === check expiration date ===
    '===================================================================================
    Sub checkExpirationDate()

        Dim currentDate As Date = Date.Now
        Dim lastRunDate As Date = My.Settings.lastRun
        Dim daysStayed As Int32 = My.Settings.expireDate.Subtract(currentDate).Days

        mainForm.menuItem_department.Enabled = False
        mainForm.menuItem_company.Enabled = False

        If lastRunDate.Subtract(currentDate).Days > 0 Then
            MsgBox("Check date and time settings!")
            mainForm.Close()
        Else
            My.Settings.lastRun = currentDate
            My.Settings.Save()
        End If

        If daysStayed > 0 Then
            Return
        Else
            MsgBox("This app has expired!")
            mainForm.Close()
        End If
    End Sub

    '===================================================================================
    '             === get file names  and dir names ===
    '===================================================================================
    Sub getNames()
        createBackup(timeStampFolder())     ' timeStampFolder() - function returns "folderName"
    End Sub
    '===================================================================================
    '             === create backup ===
    '===================================================================================
    Sub createBackup(_folderName As String)

        Dim backUpFile, foundFile As String

        '   create backup database in BackUp folder
        backUpFile = Directory.GetCurrentDirectory() & "\BackUp\" & _folderName & "\DB.ombckp"
        Try
            For Each foundFile In My.Computer.FileSystem.GetFiles _
        (mainForm.sDir, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.omdb")
                'Console.WriteLine(foundFile)
            Next
            MsgBox("Создаем резервную копию базы данных в папке BackUp", vbOKOnly + vbInformation)
            My.Computer.FileSystem.CopyFile(foundFile, backUpFile)
        Catch
        End Try
    End Sub
    '===================================================================================
    '             === extract files ===
    '===================================================================================
    Sub extractFiles()

        Using zip As ZipFile = ZipFile.Read(mainForm.sFilePath)

            zip.Password = "iSakha2836"
            zip.ExtractAll(mainForm.sDir & "\Temp", ExtractExistingFileAction.OverwriteSilently)

        End Using
        mainForm.sDir = mainForm.sDir & "\Temp"
    End Sub
    '===================================================================================
    '             === Load database ===
    '===================================================================================

    Sub load_db()
        '   Get list of database files , names of each file, list of worksheets in each file

        Dim name As String                           ' variable to get name of database file

        mainForm.fileNames = New Collection         ' collection of names of each file
        mainForm.filePath = New Collection         ' collection of full path of each file

        Dim key As Integer = 0

        mainForm.i_superPivotDict = New Dictionary(Of Integer, Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable)))
        mainForm.i_pivotTableDict = New Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable))
        mainForm.i_pivot_wsDict = New Dictionary(Of Integer, Dictionary(Of Integer, ExcelWorksheet))

        For Each foundFile As String In My.Computer.FileSystem.GetFiles _
(mainForm.sDir, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.omdb")

            '   Extract file name from full path

            mainForm.sFilePath = CStr(foundFile)         ' full path to database file

            mainForm.filePath.Add(mainForm.sFilePath)

            Dim SplitFileName_DB() As String
            SplitFileName_DB = Split(mainForm.sFilePath, "\")
            name = SplitFileName_DB(SplitFileName_DB.Count - 1)
            SplitFileName_DB = Split(name, ".")
            name = SplitFileName_DB(0)

            mainForm.fileNames.Add(name)             ' add name of each file to name collection


            '   Create collection of Excel files workSheets

            Dim ws As ExcelWorksheet
            Dim excelFile = New FileInfo(mainForm.sFilePath)
            'Console.WriteLine(mainForm.sFilePath)
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
    '             === compress files ===
    '===================================================================================
    Sub compressFiles()

        Try

            Dim sFolderPath As String = mainForm.sDir               ' Temp folder
            '           Console.WriteLine(sFolderPath)

            Dim Files() As String = IO.Directory.GetFiles(sFolderPath)

            Using zip As ZipFile = New ZipFile

                zip.Password = "iSakha2836"
                zip.Encryption = Ionic.Zip.EncryptionAlgorithm.WinZipAes256
                zip.AddDirectory(sFolderPath)
                zip.Save(Directory.GetCurrentDirectory() & "\database\DB.omdb")

            End Using

        Catch
        End Try

    End Sub

    Sub deleteTemp()
        Dim sFolderPath As String = mainForm.sDir               ' Temp folder
        Try
            My.Computer.FileSystem.DeleteDirectory(sFolderPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
        Catch
        End Try
    End Sub
    '===================================================================================
    '             === My Functions ===
    '===================================================================================
    Function timeStampFolder()

        Dim folderName, backUpFolder As String

        Dim format As String = ("yyy MM dd HH':'mm':'ss")
        Dim myDate As DateTime = DateTime.Now

        folderName = myDate.ToString(format)

        folderName = Regex.Replace(folderName, "\D", "")            ' timestamp name
        backUpFolder = Directory.GetCurrentDirectory() & "\BackUp"
        '   Create folder with timestamp name inside backUp folder
        My.Computer.FileSystem.CreateDirectory(backUpFolder & "\" & folderName)

        mainForm.sDir = Directory.GetCurrentDirectory()

        mainForm.OFD.InitialDirectory = Directory.GetCurrentDirectory()
        mainForm.OFD.Title = "Select .omdb file"
        mainForm.OFD.Filter = "Database files|*.omdb"
        '   open file using open file dialog
        If mainForm.OFD.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            mainForm.sFilePath = mainForm.OFD.FileName
        Else
            mainForm.Close()
        End If

        Dim testArray() As String = Split(mainForm.sFilePath, "\")
        Dim sFileName As String = testArray(testArray.Length - 1)
        mainForm.sDir = ""
        For i As Integer = 0 To testArray.Length - 2
            If i < testArray.Length - 2 Then
                mainForm.sDir = mainForm.sDir & testArray(i) & "\"
            Else
                mainForm.sDir = mainForm.sDir & testArray(i)
            End If
        Next i
        Return (folderName)
    End Function

    Sub loadFromBackup()
        Dim backUpFile, newDB_file As String
        newDB_file = Directory.GetCurrentDirectory() & "\database\DB_fromBackUp.omdb"
        mainForm.OFD.InitialDirectory = Directory.GetCurrentDirectory()
        mainForm.OFD.Title = "Select .ombckp file"
        mainForm.OFD.Filter = "BackUp files|*.ombckp"

        '   copy and rename file to db location
        If mainForm.OFD.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            backUpFile = mainForm.OFD.FileName
            My.Computer.FileSystem.CopyFile(backUpFile, newDB_file)
        Else
            mainForm.Close()
        End If
    End Sub

End Module
