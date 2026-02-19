Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data

Namespace Infrastructure

    Public Module ResultTableFilter

        ' Warning/Error만 남기고 나머지(None/Info)는 제거
        Public Sub KeepOnlyIssues(ByVal key As String, ByVal table As DataTable)
            If table Is Nothing Then Return
            If table.Rows Is Nothing OrElse table.Rows.Count = 0 Then Return

            For i As Integer = table.Rows.Count - 1 To 0 Step -1
                Dim st As ExcelStyleHelper.RowStatus = ExcelExportStyleRegistry.Resolve(key, table.Rows(i), table)

                Dim keep As Boolean = (st = ExcelStyleHelper.RowStatus.Warning OrElse st = ExcelStyleHelper.RowStatus.Error)
                If Not keep Then
                    table.Rows.RemoveAt(i)
                End If
            Next
        End Sub

        ' FamilyIndex 같은 보조 테이블에서 "오류가 있는 대상만" 남길 때 사용
        Public Sub KeepOnlyByNameSet(ByVal table As DataTable,
                                     ByVal nameColumn As String,
                                     ByVal keepNames As HashSet(Of String))
            If table Is Nothing Then Return
            If keepNames Is Nothing Then Return
            If table.Rows Is Nothing OrElse table.Rows.Count = 0 Then Return
            If Not table.Columns.Contains(nameColumn) Then Return

            For i As Integer = table.Rows.Count - 1 To 0 Step -1
                Dim nm As String = ""
                If table.Rows(i) IsNot Nothing AndAlso table.Rows(i)(nameColumn) IsNot Nothing Then
                    nm = Convert.ToString(table.Rows(i)(nameColumn)).Trim()
                End If
                If String.IsNullOrWhiteSpace(nm) OrElse Not keepNames.Contains(nm) Then
                    table.Rows.RemoveAt(i)
                End If
            Next
        End Sub

    End Module

End Namespace
