Friend Module Helpers

    Function RunCommand(cmdString As String) As Integer

        Using process As New Process()
            Dim startInfo As New ProcessStartInfo() With {
                .FileName = Globals.ConsolePath,
                .Arguments = cmdString,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .UseShellExecute = False,
                .CreateNoWindow = True
            }

            process.StartInfo = startInfo
            process.Start()

            Dim output As String = process.StandardOutput.ReadToEnd()
            Dim err As String = process.StandardError.ReadToEnd()

            process.WaitForExit()

            ' Optional: log or expose output somewhere
            ' Debug.WriteLine(output)
            ' Debug.WriteLine(err)

            Return process.ExitCode
        End Using

    End Function

End Module
