Imports System
Imports System.Data

Namespace Exports

    Public Module ConnectorExport

        Public Function SaveWithDialog(resultTable As DataTable) As String
            If resultTable Is Nothing Then Return String.Empty
            EnsureMessageRow(resultTable)

            Dim outPath As String = Global.KKY_Tool_Revit.Infrastructure.ExcelCore.PickAndSaveXlsx(
                "Connector Diagnostics",
                resultTable,
                "ConnectorDiagnostics.xlsx"
            )

            If String.IsNullOrWhiteSpace(outPath) Then Return outPath

            Return outPath
        End Function

        Public Sub Save(outPath As String, resultTable As DataTable)
            If String.IsNullOrWhiteSpace(outPath) Then Exit Sub
            If resultTable Is Nothing Then Exit Sub
            EnsureMessageRow(resultTable)

            Global.KKY_Tool_Revit.Infrastructure.ExcelCore.SaveXlsx(outPath, "Connector Diagnostics", resultTable)
        End Sub

        Private Sub EnsureMessageRow(table As DataTable)
            Global.KKY_Tool_Revit.Infrastructure.ExcelCore.EnsureMessageRow(table, "오류가 없습니다.")
        End Sub

    End Module

End Namespace
