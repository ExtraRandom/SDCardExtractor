Imports System.IO
Imports System

Module Module1
    Dim Desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) 'get desktop path
    Dim CopyCount As Integer = 0  'for counting files transfered
    Dim SDPicsFolder As String = Desktop + ("\SDPics") 'folder on the desktop where pictures will go
    Dim SDCard As String = ""
    Dim DCIM As String = "DCIM\"
    Sub Main()
        Console.Title = "SD Card Picture Extract"
        Console.WriteLine("SD Card Picture Extract: Version 1.3")
        FindSD()
        Console.WriteLine("SD Card Found")
        Console.WriteLine("")
        Console.WriteLine("The Pictures will be transfered to the Desktop in the folder 'SDPics'.")
        If Not My.Computer.FileSystem.DirectoryExists(SDPicsFolder) = True Then
            System.IO.Directory.CreateDirectory(SDPicsFolder)
        End If
        Dim TimeDate As String = String.Format("{0:dd.MM.yyyy HH.mm.ss}", DateTime.Now) 'for creating unique name for the folder
        Dim NewFolder As String = (SDPicsFolder & "\" & TimeDate)
        My.Computer.FileSystem.CreateDirectory(NewFolder)
        Console.WriteLine("Copying Files, Do NOT close")
        For Each File As String In My.Computer.FileSystem.GetFiles(SDCard, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories)
            Dim CheckMISC As String = File.Remove(0, 8)
            Dim FileTo As String = File.Remove(0, 7)
            Dim test() As String = Split(CheckMISC, "\"c)
            Dim FolderFileTo As String = NewFolder + FileTo
            If Not test(0) = "MISC" Then
                CopyCount = CopyCount + 1
                Console.WriteLine("#" & CopyCount & ": " & File)
                Copy(File, FolderFileTo)
            Else
                Console.WriteLine("Not Copying MISC Subfolder files")
            End If
        Next

        Console.WriteLine("")
        Console.WriteLine("Photos have been transfered. Press 'enter' to close the program.")
        Console.ReadLine()
    End Sub
    Sub FindSD()
        Dim Removable(26) As String
        Dim SDDetected As Boolean
        Dim count As Integer = 0
        Dim NotSD As Integer = 0
        Dim allDrives() As DriveInfo = DriveInfo.GetDrives
        Dim d As DriveInfo
        For Each d In allDrives 'Find all removable drives
            Console.WriteLine(d.Name)
            If d.DriveType = 2 Then '2 = removable
                count = count + 1
                Removable(count) = d.ToString
            End If
        Next
        If count = 0 Then 'if no removable found
            SDDetected = False
        Else 'if removeable was found
            For i As Integer = 0 To count
                SDCard = Removable(i) & DCIM
                Console.WriteLine(SDCard)
                If Check(SDCard) = True Then 'if (letter) contains DCIM folder
                    SDDetected = True 'SD Card found return to Main Sub
                    Exit Sub
                Else
                    NotSD = NotSD + 1
                End If
            Next
            If NotSD = count Then
                SDDetected = False
            End If
        End If
        If SDDetected = False Then 'No SD Card found, close program
            Console.WriteLine("No SD Card has been detected.")
            Console.WriteLine("Press 'Enter' to Close the Program")
            Console.ReadLine()
            End
        End If
    End Sub
    Sub Copy(ByVal sourceimg As String, ByVal desination As String)
        My.Computer.FileSystem.CopyFile(sourceimg, desination)
    End Sub
    Function Check(ByVal letter As String) As Boolean
        Dim result As Boolean
        If My.Computer.FileSystem.DirectoryExists(letter) Then
            result = True
        Else
            result = False
        End If
        Return (result)
    End Function
End Module
