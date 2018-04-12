Public Class INIFile
    Private ListSections As New Collections.Generic.Dictionary(Of String, Collections.Generic.Dictionary(Of String, String))

    Public Function AddSection(key As String) As Collections.Generic.Dictionary(Of String, String)
        Dim ValuesDictionary As New Collections.Generic.Dictionary(Of String, String)
        ListSections.Add(key, ValuesDictionary)
        Return ValuesDictionary
    End Function

    Public ReadOnly Property Count As Integer
        Get
            Return ListSections.Count
        End Get
    End Property

    Public Sub Remove(key As String)
        ListSections.Remove(key)
    End Sub

    Public ReadOnly Property Keys As String()
        Get
            Dim TKeys() As String = {}
            For Each n In ListSections.Keys
                ReDim Preserve TKeys(TKeys.Length)
                TKeys(TKeys.Length - 1) = n
            Next
            Return TKeys
        End Get
    End Property

    Public ReadOnly Property Section(key As String) As Collections.Generic.Dictionary(Of String, String)
        Get
            Return ListSections(key)
        End Get
    End Property

    Public Sub Clear()
        ListSections.Clear()
    End Sub

    Public Sub Save(FileName As String)
        Dim FileString As String = ""
        For Each SectionItem In ListSections
            Dim SectionName As String = String.Format("[{0}]", SectionItem.Key)
            FileString &= SectionName & vbNewLine
            For Each item In ListSections.Item(SectionItem.Key)
                Dim ValueData As String = String.Format("{0} = {1}", {item.Key, item.Value})
                FileString &= ValueData & vbNewLine
            Next
        Next
        IO.File.WriteAllText(FileName, FileString, Text.Encoding.GetEncoding(1251))
    End Sub

    Public Shared Function Load(FileName) As INIFile
        Dim Result As New INIFile
        Dim EndSection As Collections.Generic.Dictionary(Of String, String) = Nothing
        For Each Line In IO.File.ReadAllLines(FileName, Text.Encoding.GetEncoding(1251))
            Line = Line.Trim
            If Line.StartsWith(";") Or Line.EndsWith("//") Then
                Continue For
            ElseIf Line.StartsWith("[") And Line.EndsWith("]") Then
                EndSection = Result.AddSection(Mid(Line, 2, Line.Length - 2))
            ElseIf Line Like "*=*" Then
                Dim Values As String() = Split(Line, "=")
                EndSection.Add(Values(0).Trim, Values(1).Trim)
            End If
        Next
        Return Result
    End Function

End Class