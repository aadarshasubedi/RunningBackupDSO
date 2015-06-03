Imports System.Globalization
Imports System.IO
Imports System.Security

Module Module1

    Private sDestPath As String, sFolderPrMth As String, checkFolder As String
    ''sDestPath descrive il percorso per il salvataggio dinamico della cartella di backup
    ''sFolderPrMth descrive il percorso della cartella fonte
    ''' <summary>
    ''' used to clean up folder after backup
    ''' </summary>
    ''' <remarks></remarks>
    Sub Main()

        Try
            Dim fullTrust As New PermissionSet(Permissions.PermissionState.Unrestricted)
            fullTrust.Demand()

            'ChDrive("T")
            'ChDir("T:\Aristide_Lapa\MacrAris")

            'If Not Environ("username").ToUpper = "KWEMARIT" Then
            '    Console.WriteLine("Spiacente! non sei @ris" & vbNewLine & "Interruzione App...")
            '    Console.WriteLine(vbNewLine & vbNewLine)
            '    Console.WriteLine("Tu sei {0} <<<>>>> Buona Giornata!", Environ("username"))
            '    Console.ReadLine()
            '    Exit Sub
            'End If
            Call BaiDsoBackup()
            Call BbmiDsoBackup()
            ChDrive("C")
            Console.WriteLine("done")
            Console.ReadLine()
        Catch ex As Exception
            Console.WriteLine("Interruzione App. causa errore seguente >>> - {0} -----{1}", ex.Message, _
                                                 vbNewLine & vbNewLine & "Contattare maCR@ris...")
            Console.ReadLine()
        End Try


    End Sub

    Sub BaiDsoBackup()


        sDestPath = "\\bbmi01\Comune\Aristide_Lapa\MacrAris\Bai_Current_Month\BAI_Archive\BAI_Macro_" _
                    & sdhLastDayInMonth(Now()) _
                        & "_Backup_" & Format(Now, "yyyy-MM-dd_hh-mm-ss")

        sFolderPrMth = "\\bbmi01\Comune\Aristide_Lapa\MacrAris\Bai_Current_Month\BAI_Macro\"

        checkFolder = "BAI_Macro_" _
                & sdhLastDayInMonth(Now())
        ' '''Esegue verifca presenza file di backup creato con task scheduler tutti 3o lunedi, martedi, mercoledi del mese
        Dim Cartella_scr As String = "\\bbmi01\Comune\Aristide_Lapa\MacrAris\Bai_Current_Month\BAI_Archive\"

        If VerificaFolder(checkFolder, Cartella_scr) Then '= True Then
            Exit Sub

        Else

            My.Computer.FileSystem.CopyDirectory(sFolderPrMth, sDestPath, True)

        End If

        On Error Resume Next
        For Each foundFile In My.Computer.FileSystem.GetFiles(
            sFolderPrMth)
            Select Case My.Computer.FileSystem.GetFileInfo(foundFile).Name.ToString.ToUpper
                Case "BAI_ANAGR.XLSX"
                Case "EXTRACT_BAI_INDICICLIENTI.XLS"
                Case "BAI_AG.XLSX"
                Case "FATT_MENSILE.XLS"
                Case "BAI_LISA_RAW.XLSX"
                Case "BAI_FATT_REVOLVING_YR.XLS"
                Case "MAX_PRIVATI_AGINGBAI.XLSX"
                Case Else : My.Computer.FileSystem.DeleteFile(foundFile)

            End Select
        Next

        For Each FoundDirectory In My.Computer.FileSystem.GetDirectories(sFolderPrMth)
            My.Computer.FileSystem.DeleteDirectory(FoundDirectory, _
                                                   FileIO.DeleteDirectoryOption.DeleteAllContents)
        Next

    End Sub

    Sub BbmiDsoBackup()

        sDestPath = "\\bbmi01\Comune\Aristide_Lapa\MacrAris\Bbmi_Current_Month\BBMI_Archive\BBMI_Macro_" _
                            & sdhLastDayInMonth(Now()) _
                                    & "_Backup_" & Format(Now, "yyyy-MM-dd_hh-mm-ss")

        sFolderPrMth = "\\bbmi01\Comune\Aristide_Lapa\MacrAris\Bbmi_Current_Month\BBMI_Macro\"

        checkFolder = "BBMI_Macro_" _
                & sdhLastDayInMonth(Now())

        Dim Cartella_scr As String = "\\bbmi01\Comune\Aristide_Lapa\MacrAris\Bbmi_Current_Month\BBMI_Archive\"

        If VerificaFolder(checkFolder, Cartella_scr) Then '= True Then
            Exit Sub

        Else

            My.Computer.FileSystem.CopyDirectory(sFolderPrMth, sDestPath, True)

        End If

        On Error Resume Next

        For Each foundFile In My.Computer.FileSystem.GetFiles(sFolderPrMth)
            Select Case My.Computer.FileSystem.GetFileInfo(foundFile).Name.ToString.ToUpper
                Case "BBMI_ANAGR.XLSX"
                Case "EXTRACT_BBMI_INDICICLIENTI.XLS"
                Case "BBMI_AG.XLSX"
                Case "FATT_MENSILE.XLS"
                Case "BBMI_LISA_RAW.XLSX"
                Case "BBMI_FATT_REVOLVING_YR.XLS"
                Case "BBMI_ASSODSO.XLSX"
                Case Else : My.Computer.FileSystem.DeleteFile(foundFile)

            End Select
        Next

        For Each FoundDirectory In My.Computer.FileSystem.GetDirectories(sFolderPrMth)
            My.Computer.FileSystem.DeleteDirectory(FoundDirectory, _
                                                   FileIO.DeleteDirectoryOption.DeleteAllContents)
        Next

    End Sub

    Function VerificaFolder(folderToCheck As String, srcFolder As String) As Boolean

        For Each foundDirectory As String In
               My.Computer.FileSystem.GetDirectories(srcFolder, _
                   FileIO.SearchOption.SearchTopLevelOnly, "*" & folderToCheck & "*")
        Next
        Return True
    End Function

    Function sdhLastDayInMonth(dtmDate As Date) As String

        Dim dhLastDayInMonth As Date = DateSerial(Year(dtmDate), _
         Month(dtmDate), 0)
        sdhLastDayInMonth = Format(dhLastDayInMonth, "yyyy-MM")

    End Function

End Module
