Module Module1

    Sub Main()

        Console.WriteLine("VisualStudio Solution Boilerplate Generator")
        Console.WriteLine("デスクトップにVisualStudio Solutionを生成します")
        Console.WriteLine("----------------------------------------------------------------")
        Console.WriteLine("生成するソリューション名")

        Dim SolutionName = Console.ReadLine().Trim
        If String.IsNullOrEmpty(SolutionName) Then Exit Sub

        Dim InvalidChars = System.IO.Path.GetInvalidFileNameChars()
        If (SolutionName.IndexOfAny(InvalidChars) >= 0) Then
            Exit Sub
        End If

        Dim DesktopPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Dim SolutionPath = IO.Path.Combine(DesktopPath, SolutionName)

        '-- ソリューション内のディレクトリ
        Dim InnerDirs = New String() {"CMD", "Component", "DBInitiator", "WPF"}

        Try
            Dim DirInfo As System.IO.DirectoryInfo
            For Each InnerDir In InnerDirs
                DirInfo = System.IO.Directory.CreateDirectory(IO.Path.Combine(SolutionPath, InnerDir))
            Next
        Catch ex As Exception
            Console.WriteLine("エラー：ソリューションディレクトリが作成できませんでした")
            Exit Sub
        End Try

        '-- ソリューションファイル
        Try
            Using sw As New IO.StreamWriter($"{SolutionPath}\{SolutionName}.sln", append:=False, encoding:=System.Text.Encoding.UTF8)
                sw.WriteLine("Microsoft Visual Studio Solution File, Format Version 12.00")
                sw.WriteLine("# Visual Studio Version 16")
                sw.WriteLine("VisualStudioVersion = 16.0.33529.622")
                sw.WriteLine("MinimumVisualStudioVersion = 10.0.40219.1")
                For Each InnerDir In InnerDirs
                    sw.WriteLine($"Project(""{{2150E333-8FDC-42A3-9474-1A3956D46DE8}}"") = ""{InnerDir}"", ""{InnerDir}"", ""{Guid.NewGuid.ToString("B")}""")
                    sw.WriteLine("EndProject")
                Next
                sw.WriteLine("Global")
                sw.WriteLine($"{vbTab}GlobalSection(SolutionProperties) = preSolution")
                sw.WriteLine($"{vbTab}{vbTab}HideSolutionNode = False")
                sw.WriteLine($"{vbTab}EndGlobalSection")
                sw.WriteLine($"{vbTab}GlobalSection(ExtensibilityGlobals) = postSolution")
                sw.WriteLine($"{vbTab}{vbTab}SolutionGuid = {Guid.NewGuid.ToString("B")}")
                sw.WriteLine($"{vbTab}EndGlobalSection")
                sw.WriteLine("EndGlobal")
            End Using
        Catch ex As Exception
            Console.WriteLine("エラー：ソリューションファイルが作成できませんでした")
            Exit Sub
        End Try

        Console.WriteLine($"デスクトップに{SolutionName}が生成されました")

        '-- git設定ファイルの書き出し
        Dim assm = Reflection.Assembly.GetExecutingAssembly()
        Dim gitattributes = "VSSolution_Boilerplate..gitattributes"
        Dim gitignore = "VSSolution_Boilerplate..gitignore"

        Dim GitFiles = New String() {".gitattributes", ".gitignore"}
        For Each GitFile In GitFiles
            Using EmbeddedResourceStream = assm.GetManifestResourceStream($"VSSolution_Boilerplate.{GitFile}")

                If EmbeddedResourceStream Is Nothing Then Exit For

                Using sr As New IO.StreamReader(EmbeddedResourceStream)
                    Dim a = sr.ReadToEnd
                    Using sw As New IO.StreamWriter(IO.Path.Combine(SolutionPath, GitFile), append:=False)
                        sw.Write(a)
                    End Using
                End Using

            End Using
        Next

        Try
            'バッチファイルはSJISで書き出す
            Using sw As New IO.StreamWriter($"{SolutionPath}\ExecuteFirst.bat", append:=False, encoding:=System.Text.Encoding.GetEncoding("shift_jis"))
                sw.WriteLine("echo off")
                sw.WriteLine("git init .")
                sw.WriteLine("git commit --allow-empty -m ""InitialCommit""")
                sw.WriteLine("git add .")
                sw.WriteLine("git commit -m ""Solution Boilerplate""")
                sw.WriteLine("echo ---------------------------------------------------")
                sw.WriteLine("")
                sw.WriteLine("pause")
            End Using
        Catch ex As Exception
            Console.WriteLine("エラー：git用のバッチファイルが作成できませんでした")
            Exit Sub
        End Try

        Console.WriteLine($"git設定ファイルが生成されました")
        Threading.Thread.Sleep(5000)

    End Sub

End Module