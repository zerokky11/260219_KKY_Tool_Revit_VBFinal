Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Text
Imports WinForms = System.Windows.Forms
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI

Namespace Services

    '==================== 공용: 로드 옵션 ====================
    Friend Class FamilyLoadOptionsAllOverwrite
        Implements IFamilyLoadOptions

        Public Function OnFamilyFound(familyInUse As Boolean,
                                      ByRef overwriteParameterValues As Boolean) As Boolean _
            Implements IFamilyLoadOptions.OnFamilyFound
            overwriteParameterValues = True
            Return True
        End Function

        Public Function OnSharedFamilyFound(sharedFamily As Family,
                                            familyInUse As Boolean,
                                            ByRef source As FamilySource,
                                            ByRef overwriteParameterValues As Boolean) As Boolean _
            Implements IFamilyLoadOptions.OnSharedFamilyFound
            source = FamilySource.Family
            overwriteParameterValues = True
            Return True
        End Function
    End Class

    '==================== 공용: 실패 전처리/트랜잭션 유틸 ====================
    Friend Class SwallowWarningsOnly
        Implements IFailuresPreprocessor

        Public Function PreprocessFailures(failAcc As FailuresAccessor) As FailureProcessingResult _
            Implements IFailuresPreprocessor.PreprocessFailures

            For Each f In failAcc.GetFailureMessages()
                If f.GetSeverity() = FailureSeverity.Warning Then
                    failAcc.DeleteWarning(f)
                End If
            Next

            Return FailureProcessingResult.Continue
        End Function
    End Class

    Friend Module TxnUtil

        Public Sub WithTxn(doc As Document, name As String, action As Action)
            Using t As New Transaction(doc, name)
                t.Start()

                Dim opt = t.GetFailureHandlingOptions()
                opt.SetFailuresPreprocessor(New SwallowWarningsOnly())
                opt.SetClearAfterRollback(True)
                t.SetFailureHandlingOptions(opt)

                action.Invoke()

                t.Commit()
            End Using
        End Sub

        ' LoadFamily는 트랜잭션 밖에서 호출
        Public Sub SafeLoadFamily(sourceFamDoc As Document,
                                  targetDoc As Document,
                                  Optional label As String = "LoadFamily",
                                  Optional saveModifiedCopy As Boolean = False,
                                  Optional saveFolder As String = Nothing)

            If targetDoc Is Nothing Then Throw New ArgumentNullException(NameOf(targetDoc))

            If targetDoc.IsModifiable Then
                Throw New InvalidOperationException("Target document has an open transaction. Close it before LoadFamily.")
            End If

            If saveModifiedCopy AndAlso sourceFamDoc IsNot Nothing AndAlso sourceFamDoc.IsModified Then
                SaveModifiedFamilyCopy(sourceFamDoc, saveFolder)
            End If

            Dim opts As New FamilyLoadOptionsAllOverwrite()
            sourceFamDoc.LoadFamily(targetDoc, opts)

            Try : targetDoc.Regenerate() : Catch : End Try
        End Sub

        Private Sub SaveModifiedFamilyCopy(sourceFamDoc As Document, saveFolder As String)
            If sourceFamDoc Is Nothing OrElse Not sourceFamDoc.IsModified Then Return

            Dim normalizedFolder As String = If(saveFolder, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedFolder) Then
                Throw New InvalidOperationException("수정된 패밀리(.rfa) 저장 폴더가 지정되지 않았습니다.")
            End If

            Directory.CreateDirectory(normalizedFolder)

            Dim baseName As String = Path.GetFileNameWithoutExtension(If(sourceFamDoc.Title, String.Empty))
            If String.IsNullOrWhiteSpace(baseName) Then
                baseName = "ModifiedFamily"
            End If

            For Each ch As Char In Path.GetInvalidFileNameChars()
                baseName = baseName.Replace(ch, "_"c)
            Next

            Dim savePath As String = Path.Combine(normalizedFolder, baseName & ".rfa")
            Dim options As New SaveAsOptions()
            options.OverwriteExistingFile = True

            sourceFamDoc.SaveAs(savePath, options)
        End Sub

    End Module

    '==================== 공유 패밀리 판정 ====================
    Friend Module SharedFamilyHelper

        Private ReadOnly _cache As New Dictionary(Of Integer, Boolean)()

        ''' <summary>
        ''' 해당 패밀리의 "공유 패밀리(Shared)" 체크 여부를 ownerDocCanEdit 기준으로 판정.
        ''' </summary>
        Public Function IsFamilyShared(ownerDocCanEdit As Document, fam As Family) As Boolean
            If fam Is Nothing Then Return False

            Dim key As Integer = fam.Id.IntValue()
            Dim cached As Boolean
            If _cache.TryGetValue(key, cached) Then Return cached

            Dim fdoc As Document = Nothing
            Try
                fdoc = ownerDocCanEdit.EditFamily(fam)
                Dim p As Parameter = fdoc.OwnerFamily.Parameter(BuiltInParameter.FAMILY_SHARED)
                Dim isSharedFlag As Boolean = (p IsNot Nothing AndAlso p.AsInteger() = 1)
                _cache(key) = isSharedFlag
                Return isSharedFlag
            Catch
                _cache(key) = False
                Return False
            Finally
                If fdoc IsNot Nothing Then
                    Try : fdoc.Close(False) : Catch : End Try
                End If
            End Try
        End Function

    End Module

    '==================== 파라미터 타입 문자열 헬퍼 ====================
    Friend Module SharedParamTypeHelper

        Friend Function GetParamTypeString(ed As ExternalDefinition) As String
            If ed Is Nothing Then Return String.Empty

#If REVIT2019 Or REVIT2021 Then
            Return ed.ParameterType.ToString()
#ElseIf REVIT2023 Or REVIT2025 Then
            Try
                Dim dataType As ForgeTypeId = ed.GetDataType()
                If dataType Is Nothing Then Return String.Empty

                Try
                    Return LabelUtils.GetLabelForSpec(dataType)
                Catch
                End Try

                Try
                    Return dataType.TypeId
                Catch
                End Try

                Return dataType.ToString()
            Catch
                Return String.Empty
            End Try
#End If
            Return String.Empty
        End Function

    End Module

    '==================== 파라미터 선택 폼 ====================
    Friend Class FormSharedParamPicker
        Inherits WinForms.Form

        Private ReadOnly _app As Application
        Private _defFile As DefinitionFile

        Private txtSearch As WinForms.TextBox
        Private lstGroups As WinForms.ListBox
        Private lvParams As WinForms.ListView
        Private btnOK As WinForms.Button
        Private btnCancel As WinForms.Button
        Private chkExcludeDummy As WinForms.CheckBox
        Private cboTargetGroup As WinForms.ComboBox
        Private rbInstance As WinForms.RadioButton
        Private rbType As WinForms.RadioButton

        Public ReadOnly Property SelectedExternalDefinitions As New List(Of ExternalDefinition)
        Public ReadOnly Property ExcludeDummy As Boolean
            Get
                Return chkExcludeDummy IsNot Nothing AndAlso chkExcludeDummy.Checked
            End Get
        End Property
        Public ReadOnly Property SelectedGroupPG As BuiltInParameterGroup
            Get
                Dim item = TryCast(cboTargetGroup.SelectedItem, GroupItem)
                If item Is Nothing Then Return BuiltInParameterGroup.PG_IDENTITY_DATA
                Return item.PG
            End Get
        End Property
        Public ReadOnly Property SelectedIsInstance As Boolean
            Get
                If rbInstance Is Nothing Then Return True
                Return rbInstance.Checked
            End Get
        End Property

        Private Class GroupItem
            Public ReadOnly PG As BuiltInParameterGroup
            Public ReadOnly Name As String
            Public Sub New(pg As BuiltInParameterGroup)
                Me.PG = pg : Me.Name = pg.ToString()
            End Sub
            Public Overrides Function ToString() As String
                Return Name
            End Function
        End Class

        Public Sub New(app As Application)
            _app = app
            Me.Text = "Select Shared Parameter(s) – Multi-select"
            Me.StartPosition = WinForms.FormStartPosition.CenterScreen
            Me.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
            Me.MaximizeBox = False : Me.MinimizeBox = False
            Me.Width = 920 : Me.Height = 700
            BuildUI()
            LoadDefinitions()
            LoadGroupCombo()
        End Sub

        Private Sub BuildUI()
            Dim lblSearch As New WinForms.Label() With {.Text = "Search:", .Left = 12, .Top = 14, .AutoSize = True}
            txtSearch = New WinForms.TextBox() With {.Left = 70, .Top = 10, .Width = 820}
            AddHandler txtSearch.TextChanged, AddressOf OnSearchChanged

            Dim lblGroups As New WinForms.Label() With {
                .Text = "Shared Parameter Groups (Multi-select / (All Groups) available)",
                .Left = 12, .Top = 44, .AutoSize = True
            }
            lstGroups = New WinForms.ListBox() With {
                .Left = 12,
                .Top = 64,
                .Width = 320,
                .Height = 520,
                .SelectionMode = WinForms.SelectionMode.MultiExtended
            }
            AddHandler lstGroups.SelectedIndexChanged, AddressOf OnGroupChanged

            Dim lblParams As New WinForms.Label() With {
                .Text = "Parameters — Name / Type / Visible / Group",
                .Left = 344,
                .Top = 44,
                .AutoSize = True
            }
            lvParams = New WinForms.ListView() With {
                .Left = 344,
                .Top = 64,
                .Width = 556,
                .Height = 520,
                .View = WinForms.View.Details,
                .FullRowSelect = True,
                .MultiSelect = True
            }
            lvParams.Columns.Add("Name", 240)
            lvParams.Columns.Add("Type", 100)
            lvParams.Columns.Add("Visible", 70)
            lvParams.Columns.Add("Group", 130)
            AddHandler lvParams.DoubleClick, AddressOf OnParamDoubleClick

            chkExcludeDummy = New WinForms.CheckBox() With {
                .Text = "하위 패밀리 이름에 'Dummy' 포함 시 제외 (기본)",
                .Left = 12,
                .Top = 592,
                .Width = 480,
                .Checked = True
            }

            Dim lblTarget As New WinForms.Label() With {
                .Text = "추가할 파라미터 그룹:",
                .Left = 344,
                .Top = 592,
                .AutoSize = True
            }
            cboTargetGroup = New WinForms.ComboBox() With {
                .Left = 480,
                .Top = 588,
                .Width = 220,
                .DropDownStyle = WinForms.ComboBoxStyle.DropDownList
            }

            rbInstance = New WinForms.RadioButton() With {
                .Text = "인스턴스",
                .Left = 720,
                .Top = 588,
                .AutoSize = True,
                .Checked = True
            }
            rbType = New WinForms.RadioButton() With {
                .Text = "타입",
                .Left = 800,
                .Top = 588,
                .AutoSize = True
            }

            btnOK = New WinForms.Button() With {.Text = "OK", .Left = 716, .Top = 620, .Width = 85}
            AddHandler btnOK.Click, AddressOf OnOK
            btnCancel = New WinForms.Button() With {.Text = "Cancel", .Left = 815, .Top = 620, .Width = 85}
            AddHandler btnCancel.Click, AddressOf OnCancel

            Controls.AddRange(New WinForms.Control() {
                lblSearch, txtSearch,
                lblGroups, lstGroups,
                lvParams,
                chkExcludeDummy,
                lblTarget, cboTargetGroup,
                rbInstance, rbType,
                btnOK, btnCancel
            })
        End Sub

        Private Sub LoadDefinitions()
            SelectedExternalDefinitions.Clear()
            lvParams.Items.Clear()
            lstGroups.Items.Clear()

            If String.IsNullOrEmpty(_app.SharedParametersFilename) OrElse Not File.Exists(_app.SharedParametersFilename) Then
                Throw New InvalidOperationException("Revit 옵션에서 Shared Parameters 파일을 지정해 주세요.")
            End If

            _defFile = _app.OpenSharedParameterFile()
            If _defFile Is Nothing Then Throw New InvalidOperationException("Shared Parameters 파일을 열 수 없습니다.")

            lstGroups.Items.Add("(All Groups)")
            For Each g As DefinitionGroup In _defFile.Groups
                lstGroups.Items.Add(g.Name)
            Next

            If lstGroups.Items.Count > 0 Then
                lstGroups.SelectedIndices.Clear()
                lstGroups.SelectedIndex = 0
            End If

            PopulateParams()
        End Sub

        Private Sub LoadGroupCombo()
            cboTargetGroup.Items.Clear()
            Dim preferred As BuiltInParameterGroup() = {
                BuiltInParameterGroup.PG_TEXT,
                BuiltInParameterGroup.PG_IDENTITY_DATA,
                BuiltInParameterGroup.PG_DATA,
                BuiltInParameterGroup.PG_CONSTRAINTS
            }
            Dim added As New HashSet(Of BuiltInParameterGroup)()
            For Each pg In preferred
                cboTargetGroup.Items.Add(New GroupItem(pg))
                added.Add(pg)
            Next

            For Each pg As BuiltInParameterGroup In [Enum].GetValues(GetType(BuiltInParameterGroup))
                If Not added.Contains(pg) Then
                    cboTargetGroup.Items.Add(New GroupItem(pg))
                End If
            Next

            Dim defIndex As Integer = -1
            For i = 0 To cboTargetGroup.Items.Count - 1
                Dim gi = TryCast(cboTargetGroup.Items(i), GroupItem)
                If gi IsNot Nothing AndAlso gi.PG = BuiltInParameterGroup.PG_TEXT Then
                    defIndex = i
                    Exit For
                End If
            Next

            If defIndex < 0 Then defIndex = 0
            cboTargetGroup.SelectedIndex = defIndex
        End Sub

        Private Sub OnGroupChanged(sender As Object, e As EventArgs)
            PopulateParams()
        End Sub

        Private Sub OnSearchChanged(sender As Object, e As EventArgs)
            PopulateParams()
        End Sub

        Private Sub OnParamDoubleClick(sender As Object, e As EventArgs)
            OnOK(sender, e)
        End Sub

        Private Sub OnOK(sender As Object, e As EventArgs)
            SelectedExternalDefinitions.Clear()
            If lvParams.SelectedItems.Count = 0 Then
                WinForms.MessageBox.Show("하나 이상의 파라미터를 선택하세요.", "KKY", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)
                Return
            End If

            For Each it As WinForms.ListViewItem In lvParams.SelectedItems
                Dim ed = TryCast(it.Tag, ExternalDefinition)
                If ed IsNot Nothing Then
                    SelectedExternalDefinitions.Add(ed)
                End If
            Next

            Me.DialogResult = WinForms.DialogResult.OK
        End Sub

        Private Sub OnCancel(sender As Object, e As EventArgs)
            Me.DialogResult = WinForms.DialogResult.Cancel
        End Sub

        Private Sub PopulateParams()
            lvParams.Items.Clear()
            If _defFile Is Nothing Then Return

            Dim selectedGroupNames As New List(Of String)()
            For Each idx As Integer In lstGroups.SelectedIndices
                selectedGroupNames.Add(lstGroups.Items(idx).ToString())
            Next

            Dim search As String = If(txtSearch.Text, String.Empty).Trim()

            Dim useAll As Boolean =
                (selectedGroupNames.Count = 0) OrElse
                (selectedGroupNames.Count = 1 AndAlso String.Equals(selectedGroupNames(0), "(All Groups)", StringComparison.OrdinalIgnoreCase))

            Dim groupsToShow As IEnumerable(Of DefinitionGroup)
            If useAll Then
                groupsToShow = _defFile.Groups
            Else
                Dim picked = New HashSet(Of String)(
                    selectedGroupNames.Where(Function(n) Not String.Equals(n, "(All Groups)", StringComparison.OrdinalIgnoreCase)),
                    StringComparer.OrdinalIgnoreCase)
                groupsToShow = _defFile.Groups.Cast(Of DefinitionGroup)().Where(Function(g) picked.Contains(g.Name))
            End If

            For Each g In groupsToShow
                For Each d As Definition In g.Definitions
                    Dim ed = TryCast(d, ExternalDefinition)
                    If ed Is Nothing Then Continue For

                    If Not String.IsNullOrEmpty(search) Then
                        Dim okName As Boolean = d.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                        Dim okGroup As Boolean = g.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                        If Not (okName OrElse okGroup) Then Continue For
                    End If

                    Dim lvi As New WinForms.ListViewItem(d.Name)
                    lvi.SubItems.Add(SharedParamTypeHelper.GetParamTypeString(ed))
                    lvi.SubItems.Add(If(ed.Visible, "Yes", "No"))
                    lvi.SubItems.Add(g.Name)
                    lvi.Tag = ed
                    lvParams.Items.Add(lvi)
                Next
            Next

            If lvParams.Items.Count > 0 Then
                lvParams.Items(0).Selected = True
            End If
        End Sub
    End Class

    '==================== 결과 폼 (엑셀 내보내기 포함) ====================
    Friend Class FormPropagateReport
        Inherits WinForms.Form

        Private Class RowInfo
            Public Property Kind As String
            Public Property Family As String
            Public Property Detail As String
        End Class

        Private ReadOnly _report As String
        Private ReadOnly _rows As List(Of RowInfo)

        Private txtReport As WinForms.TextBox
        Private lv As WinForms.ListView
        Private btnClose As WinForms.Button
        Private btnCsv As WinForms.Button

        Public Sub New(report As String,
                       scanFails As IEnumerable(Of String),
                       skips As IEnumerable(Of String),
                       parentFails As IEnumerable(Of String),
                       childFails As IEnumerable(Of String))
            _report = report
            _rows = New List(Of RowInfo)()
            BuildRows(scanFails, "ScanFail")
            BuildRows(skips, "Skip")
            BuildRows(parentFails, "Error")
            BuildRows(childFails, "ChildError")
            BuildUI()
        End Sub

        Private Sub BuildRows(items As IEnumerable(Of String), kind As String)
            If items Is Nothing Then Return
            For Each s In items
                If String.IsNullOrWhiteSpace(s) Then Continue For
                Dim fam As String = s
                Dim detail As String = String.Empty
                Dim parts = s.Split(New Char() {":"c}, 2)
                If parts.Length = 2 Then
                    fam = parts(0).Trim()
                    detail = parts(1).Trim()
                End If
                _rows.Add(New RowInfo With {.Kind = kind, .Family = fam, .Detail = detail})
            Next
        End Sub

        Private Sub BuildUI()
            Me.Text = "KKY Param Propagator - Report"
            Me.Width = 900
            Me.Height = 720
            Me.StartPosition = WinForms.FormStartPosition.CenterScreen

            txtReport = New WinForms.TextBox() With {
                .Left = 10,
                .Top = 10,
                .Width = 860,
                .Height = 260,
                .Multiline = True,
                .ScrollBars = WinForms.ScrollBars.Vertical,
                .ReadOnly = True,
                .Font = New Drawing.Font("Consolas", 9.0F),
                .Text = _report
            }

            lv = New WinForms.ListView() With {
                .Left = 10,
                .Top = 280,
                .Width = 860,
                .Height = 340,
                .View = WinForms.View.Details,
                .FullRowSelect = True
            }
            lv.Columns.Add("Type", 100)
            lv.Columns.Add("Family", 400)
            lv.Columns.Add("Detail", 340)

            For Each r In _rows
                Dim item As New WinForms.ListViewItem(r.Kind)
                item.SubItems.Add(r.Family)
                item.SubItems.Add(r.Detail)
                lv.Items.Add(item)
            Next

            btnClose = New WinForms.Button() With {.Text = "닫기", .Width = 100, .Left = 770, .Top = 630}
            AddHandler btnClose.Click, Sub() Me.Close()

            btnCsv = New WinForms.Button() With {.Text = "엑셀 내보내기", .Width = 120, .Left = 640, .Top = 630}
            AddHandler btnCsv.Click, AddressOf OnExportCsv

            Controls.AddRange(New WinForms.Control() {txtReport, lv, btnCsv, btnClose})
        End Sub

        Private Sub OnExportCsv(sender As Object, e As EventArgs)
            Using dlg As New WinForms.SaveFileDialog()
                dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                dlg.FileName = "KKY_ParamPropagator_Report.xlsx"
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return

                Try
                    Dim savePath = dlg.FileName
                    If Not String.Equals(Path.GetExtension(savePath), ".xlsx", StringComparison.OrdinalIgnoreCase) Then
                        savePath = Path.ChangeExtension(savePath, ".xlsx")
                    End If

                    Dim dt As New DataTable("SharedParamPropagate")
                    dt.Columns.Add("Type")
                    dt.Columns.Add("Family")
                    dt.Columns.Add("Detail")

                    For Each r In _rows
                        Dim row = dt.NewRow()
                        row("Type") = If(r.Kind, String.Empty)
                        row("Family") = If(r.Family, String.Empty)
                        row("Detail") = If(r.Detail, String.Empty)
                        dt.Rows.Add(row)
                    Next

                    Infrastructure.ExcelCore.SaveXlsx(savePath, "Results", dt, autoFit:=True, sheetKey:="paramprop", exportKind:="paramprop")
                    WinForms.MessageBox.Show("엑셀 내보내기가 완료되었습니다.", "KKY Param Propagator")
                Catch ex As Exception
                    WinForms.MessageBox.Show($"엑셀 내보내기 실패: {ex.Message}", "KKY Param Propagator")
                End Try
            End Using
        End Sub
    End Class

    '==================== 메인 서비스 ====================
    Public Class ParamPropagateService
        Private Shared _saveModifiedFamilies As Boolean = False
        Private Shared _modifiedFamilyOutputFolder As String = String.Empty

        Public Enum RunStatus
            Succeeded
            Cancelled
            Failed
        End Enum

        Public Class SharedParamDefinitionDto
            Public Property GroupName As String
            Public Property Name As String
            Public Property ParamType As String
            Public Property Visible As Boolean
        End Class

        Public Class ParameterGroupOption
            Public Property Id As Integer
            Public Property Key As String
            Public Property Name As String
        End Class

        Public Class SharedParamListResponse
            Public Property Ok As Boolean
            Public Property Message As String
            Public Property Definitions As List(Of SharedParamDefinitionDto)
            Public Property TargetGroups As List(Of ParameterGroupOption)
        End Class

        Public Class SharedParamRunRequest
            Public Property ParamNames As List(Of String)
            Public Property ParamGuids As List(Of String)
            Public Property TargetGroup As Integer
            Public Property TargetGroupKey As String
            Public Property IsInstance As Boolean
            Public Property ExcludeDummy As Boolean
            Public Property IncludeStandaloneFamilies As Boolean
            Public Property SaveModifiedFamilies As Boolean
            Public Property ModifiedFamilyOutputFolder As String

            Public Shared Function FromPayload(payload As Object) As SharedParamRunRequest
                Dim req As New SharedParamRunRequest With {
                    .ParamNames = New List(Of String)(),
                    .ParamGuids = New List(Of String)(),
                    .TargetGroup = CInt(BuiltInParameterGroup.PG_TEXT),
                    .TargetGroupKey = String.Empty,
                    .IsInstance = True,
                    .ExcludeDummy = True,
                    .IncludeStandaloneFamilies = False,
                    .SaveModifiedFamilies = False,
                    .ModifiedFamilyOutputFolder = String.Empty
                }

                Try
                    If payload Is Nothing Then Return req

                    Dim namesObj = ReadProp(payload, "paramNames")
                    If TypeOf namesObj Is IEnumerable Then
                        For Each n In CType(namesObj, IEnumerable)
                            If n Is Nothing Then Continue For
                            Dim raw As String = n.ToString()
                            If String.IsNullOrWhiteSpace(raw) Then Continue For

                            ' Allow both ["A","B"] and ["A,B"] input styles
                            For Each part As String In raw.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries)
                                Dim nm As String = NormalizeParamName(part)
                                If Not String.IsNullOrWhiteSpace(nm) Then req.ParamNames.Add(nm)
                            Next
                        Next
                    End If

                    Dim guidsObj = ReadProp(payload, "paramGuids")
                    If TypeOf guidsObj Is IEnumerable Then
                        For Each g In CType(guidsObj, IEnumerable)
                            If g Is Nothing Then Continue For
                            Dim gs As String = NormalizeParamName(g.ToString())
                            If Not String.IsNullOrWhiteSpace(gs) Then req.ParamGuids.Add(gs)
                        Next
                    End If

                    Dim gObj = ReadProp(payload, "group")
                    If gObj IsNot Nothing Then req.TargetGroup = Convert.ToInt32(gObj)

                    Dim gKeyObj = ReadProp(payload, "groupKey")
                    If gKeyObj Is Nothing Then gKeyObj = ReadProp(payload, "targetGroupKey")
                    If gKeyObj IsNot Nothing Then req.TargetGroupKey = NormalizeParamName(gKeyObj.ToString())

                    Dim instObj = ReadProp(payload, "isInstance")
                    If instObj IsNot Nothing Then req.IsInstance = Convert.ToBoolean(instObj)

                    Dim dummyObj = ReadProp(payload, "excludeDummy")
                    If dummyObj IsNot Nothing Then req.ExcludeDummy = Convert.ToBoolean(dummyObj)

                    Dim standaloneObj = ReadProp(payload, "includeStandaloneFamilies")
                    If standaloneObj IsNot Nothing Then req.IncludeStandaloneFamilies = Convert.ToBoolean(standaloneObj)

                    Dim saveObj = ReadProp(payload, "saveModifiedFamilies")
                    If saveObj IsNot Nothing Then req.SaveModifiedFamilies = Convert.ToBoolean(saveObj)

                    Dim folderObj = ReadProp(payload, "modifiedFamilyOutputFolder")
                    If folderObj IsNot Nothing Then req.ModifiedFamilyOutputFolder = NormalizeFolderPath(folderObj.ToString())
                Catch
                End Try

                ' De-duplicate
                req.ParamNames = req.ParamNames.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
                req.ParamGuids = req.ParamGuids.Distinct(StringComparer.OrdinalIgnoreCase).ToList()

                Return req
            End Function

            Private Shared Function ReadProp(payload As Object, name As String) As Object
                If payload Is Nothing OrElse String.IsNullOrEmpty(name) Then Return Nothing

                Try
                    Dim t = payload.GetType()
                    Dim pi = t.GetProperty(name)
                    If pi IsNot Nothing Then Return pi.GetValue(payload)

                    Dim fi = t.GetField(name)
                    If fi IsNot Nothing Then Return fi.GetValue(payload)

                    Dim dict = TryCast(payload, IDictionary)
                    If dict IsNot Nothing AndAlso dict.Contains(name) Then Return dict(name)
                Catch
                End Try

                Return Nothing
            End Function
        End Class

        Public Class SharedParamDetailRow
            Public Property Kind As String
            Public Property Family As String
            Public Property Detail As String
        End Class

        Public Class SharedParamRunResult
            Public Property Status As RunStatus
            Public Property Message As String
            Public Property Report As String
            Public Property Details As List(Of SharedParamDetailRow)
        End Class

        '==================== 목록 제공 ====================
        Public Shared Function GetSharedParameterDefinitions(app As UIApplication) As SharedParamListResponse
            Dim res As New SharedParamListResponse With {
                .Ok = False,
                .Message = Nothing,
                .Definitions = New List(Of SharedParamDefinitionDto)(),
                .TargetGroups = BuildGroupOptions()
            }

            Try
                If app Is Nothing OrElse app.Application Is Nothing Then
                    res.Message = "UIApplication 이 없습니다."
                    Return res
                End If

                Dim sharedPath As String = app.Application.SharedParametersFilename
                If String.IsNullOrEmpty(sharedPath) OrElse Not File.Exists(sharedPath) Then
                    res.Message = "공유 파라미터 파일 먼저 지정해 주세요."
                    Return res
                End If

                Dim defFile = app.Application.OpenSharedParameterFile()
                If defFile Is Nothing Then
                    res.Message = "공유 파라미터 파일을 열 수 없습니다."
                    Return res
                End If

                For Each g As DefinitionGroup In defFile.Groups
                    For Each d As Definition In g.Definitions
                        Dim ed = TryCast(d, ExternalDefinition)
                        If ed Is Nothing Then Continue For
                        res.Definitions.Add(New SharedParamDefinitionDto With {
                            .GroupName = g.Name,
                            .Name = ed.Name,
                            .ParamType = SharedParamTypeHelper.GetParamTypeString(ed),
                            .Visible = ed.Visible
                        })
                    Next
                Next

                res.Ok = True
            Catch ex As Exception
                res.Ok = False
                res.Message = ex.Message
            End Try

            Return res
        End Function

        '==================== 실행 엔트리 ====================
        Public Shared Function Run(app As UIApplication, request As SharedParamRunRequest, Optional progress As Action(Of String, Double, Integer, Integer, String, String) = Nothing) As SharedParamRunResult
            Dim result As New SharedParamRunResult With {
                .Status = RunStatus.Failed,
                .Details = New List(Of SharedParamDetailRow)()
            }

            Try
                If progress IsNot Nothing Then progress("start", 0.0R, 0, 0, "시작", "")
            Catch
            End Try

            If app Is Nothing Then
                result.Message = "UIApplication 이 없습니다."
                Return result
            End If

            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing OrElse uiDoc.Document Is Nothing Then
                result.Message = "활성 문서가 없습니다."
                Return result
            End If

            Dim doc As Document = uiDoc.Document
            If doc.IsFamilyDocument Then
                result.Message = "프로젝트 문서에서 실행하세요."
                Return result
            End If

            Dim sharedPath As String = app.Application.SharedParametersFilename
            If String.IsNullOrEmpty(sharedPath) OrElse Not File.Exists(sharedPath) Then
                result.Message = "공유 파라미터 파일 먼저 지정해 주세요."
                Return result
            End If

            If request Is Nothing OrElse request.ParamNames Is Nothing OrElse request.ParamNames.Count = 0 Then
                result.Message = "선택된 공유 파라미터가 없습니다."
                result.Status = RunStatus.Cancelled
                Return result
            End If

            If request.SaveModifiedFamilies Then
                request.ModifiedFamilyOutputFolder = NormalizeFolderPath(request.ModifiedFamilyOutputFolder)
                If String.IsNullOrWhiteSpace(request.ModifiedFamilyOutputFolder) Then
                    result.Message = "수정된 패밀리(.rfa)를 저장할 폴더를 선택해 주세요."
                    result.Status = RunStatus.Cancelled
                    Return result
                End If

                Try
                    Directory.CreateDirectory(request.ModifiedFamilyOutputFolder)
                Catch ex As Exception
                    result.Message = "수정된 패밀리(.rfa) 저장 폴더를 준비하지 못했습니다: " & ex.Message
                    Return result
                End Try
            End If

            _saveModifiedFamilies = If(request IsNot Nothing AndAlso request.SaveModifiedFamilies, True, False)
            _modifiedFamilyOutputFolder = If(request Is Nothing, String.Empty, request.ModifiedFamilyOutputFolder)

            Dim chosenPG As BuiltInParameterGroup = BuiltInParameterGroup.PG_TEXT
#If REVIT2025 Then
            Try
                Dim serializedGroup As Object = If(String.IsNullOrWhiteSpace(request.TargetGroupKey),
                                                   CType(request.TargetGroup, Object),
                                                   CType(request.TargetGroupKey, Object))
                chosenPG = BuiltInParameterGroupCompat.FromSerializedValue(serializedGroup, BuiltInParameterGroup.PG_TEXT)
            Catch
                chosenPG = BuiltInParameterGroup.PG_TEXT
            End Try
#Else
            Try
                chosenPG = CType(request.TargetGroup, BuiltInParameterGroup)
            Catch
                chosenPG = BuiltInParameterGroup.PG_TEXT
            End Try
#End If

            Dim extDefs As List(Of ExternalDefinition) = ResolveDefinitions(app.Application, request.ParamNames, request.ParamGuids)
            If extDefs Is Nothing OrElse extDefs.Count = 0 Then
                result.Message = "선택한 공유 파라미터를 Shared Parameters 파일에서 찾을 수 없습니다."
                Return result
            End If

            Dim status As RunStatus =
                ExecuteCore(doc, extDefs, request.ParamNames, request.ExcludeDummy, request.IncludeStandaloneFamilies, chosenPG, request.IsInstance, result, progress)

            result.Status = status
            If String.IsNullOrEmpty(result.Message) Then
                result.Message = If(status = RunStatus.Succeeded,
                                    "공유 파라미터 연동을 완료했습니다.",
                                    "공유 파라미터 연동에 실패했습니다.")
            End If

            Try
                If progress IsNot Nothing Then progress("done", 1.0R, 1, 1, result.Message, "")
            Catch
            End Try

            Return result
        End Function

        '==================== 결과를 엑셀로 ====================
        Public Shared Function ExportResultToExcel(result As SharedParamRunResult, Optional autoFit As Boolean = False, Optional exportLocale As String = "ko") As String
            If result Is Nothing Then Return String.Empty

            Dim hasDetails As Boolean = (result.Details IsNot Nothing AndAlso result.Details.Count > 0)
            Dim hasReport As Boolean = Not String.IsNullOrWhiteSpace(result.Report)
            If Not hasDetails AndAlso Not hasReport Then Return String.Empty

            Dim defaultName As String = $"ParamProp_{Date.Now:yyMMdd_HHmmss}.xlsx"

            Using sfd As New WinForms.SaveFileDialog()
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                sfd.FileName = defaultName
                If sfd.ShowDialog() <> WinForms.DialogResult.OK Then Return String.Empty

                Dim dt As New DataTable("SharedParamPropagate")
                dt.Columns.Add("Type")
                dt.Columns.Add("Family")
                dt.Columns.Add("Detail")

                If hasReport Then
                    For Each line As String In result.Report.Replace(vbCrLf, vbLf).Split({vbLf}, StringSplitOptions.None)
                        If String.IsNullOrWhiteSpace(line) Then Continue For
                        Dim row = dt.NewRow()
                        row("Type") = "Report"
                        row("Family") = String.Empty
                        row("Detail") = line.Trim()
                        dt.Rows.Add(row)
                    Next
                End If

                If hasDetails Then
                    For Each r In result.Details
                        Dim row = dt.NewRow()
                        row("Type") = r.Kind
                        row("Family") = r.Family
                        row("Detail") = r.Detail
                        dt.Rows.Add(row)
                    Next
                End If

                Infrastructure.ExcelCore.SaveXlsx(sfd.FileName, "Results", dt, autoFit, sheetKey:="paramprop", exportKind:="paramprop", exportLocale:=exportLocale)
                Return sfd.FileName
            End Using
        End Function

        '==================== 핵심 실행 (이름 기반 계층 + 역순 처리) ====================
        Private Shared Function ExecuteCore(doc As Document,
                                            extDefs As List(Of ExternalDefinition),
                                            paramNames As List(Of String),
                                            excludeDummy As Boolean,
                                            includeStandaloneFamilies As Boolean,
                                            chosenPG As BuiltInParameterGroup,
                                            chosenIsInstance As Boolean,
                                            result As SharedParamRunResult,
                                            Optional progress As Action(Of String, Double, Integer, Integer, String, String) = Nothing) As RunStatus

            ' ---- progress helper (Hub: paramprop:progress) ----
            Dim lastTick As Long = 0
            Dim sw As System.Diagnostics.Stopwatch = Nothing
            Try
                If progress IsNot Nothing Then
                    sw = System.Diagnostics.Stopwatch.StartNew()
                End If
            Catch
                sw = Nothing
            End Try

            Dim report As Action(Of String, Double, Integer, Integer, String, String) =
                Sub(phase As String, phaseProgress As Double, current As Integer, total As Integer, message As String, target As String)
                    If progress Is Nothing Then Return
                    Try
                        Dim doSend As Boolean = True
                        If sw IsNot Nothing Then
                            Dim nowMs As Long = sw.ElapsedMilliseconds
                            If (nowMs - lastTick) < 80 AndAlso current <> total Then
                                doSend = False
                            Else
                                lastTick = nowMs
                            End If
                        End If
                        If doSend Then progress(phase, phaseProgress, current, total, message, target)
                    Catch
                    End Try
                End Sub

            ' 1. 편집 가능한 모든 패밀리 수집 (프로젝트 문서 기준)
            Dim allEditable As List(Of Family) =
                New FilteredElementCollector(doc).
                OfClass(GetType(Family)).
                Cast(Of Family)().
                Where(Function(f) f.IsEditable).
                ToList()

            Dim totalEditableCount As Integer = allEditable.Count

            report("scan", 0.0R, 0, totalEditableCount, "패밀리 스캔", "")

            ' 이름 → Family 매핑 (Id 는 Doc별이라 안씀)
            Dim nameToFamily As New Dictionary(Of String, Family)(StringComparer.OrdinalIgnoreCase)
            For Each f In allEditable
                If f Is Nothing OrElse f.FamilyCategory Is Nothing Then Continue For
                If IsAnnotationFamily(f) Then Continue For

                If Not nameToFamily.ContainsKey(f.Name) Then
                    nameToFamily.Add(f.Name, f)
                End If
            Next

            ' 부모이름 → 자식이름 그래프 (공유 체크된 하위만)
            Dim parentToChildren As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
            Dim allTargetNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim childNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Dim dummyExcludedCount As Integer = 0
            Dim scanFails As New List(Of String)()

            '----- 1차 스캔: 그래프 구성 -----
            Dim scanIdx As Integer = 0
            For Each f In allEditable
                scanIdx += 1
                report("scan", If(totalEditableCount <= 0, 0.0R, CDbl(scanIdx) / CDbl(totalEditableCount)), scanIdx, totalEditableCount, "스캔 중...", If(f Is Nothing, "", f.Name))
                If f.FamilyCategory Is Nothing Then Continue For
                If IsAnnotationFamily(f) Then Continue For

                Dim hostDoc As Document = Nothing
                Try
                    hostDoc = doc.EditFamily(f)

                    Dim hostName As String = f.Name
                    Dim hostFm As FamilyManager = Nothing
                    Try : hostFm = hostDoc.FamilyManager : Catch : hostFm = Nothing : End Try

                    Dim insts = New FilteredElementCollector(hostDoc).
                                OfClass(GetType(FamilyInstance)).
                                Cast(Of FamilyInstance)()

                    For Each fi As FamilyInstance In insts
                        Dim childNamesForInstance = CollectChildFamiliesForGraph(hostDoc, hostFm, fi, excludeDummy, dummyExcludedCount)
                        If childNamesForInstance Is Nothing OrElse childNamesForInstance.Count = 0 Then Continue For

                        Dim setChildren As HashSet(Of String) = Nothing
                        If Not parentToChildren.TryGetValue(hostName, setChildren) Then
                            setChildren = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                            parentToChildren(hostName) = setChildren
                        End If

                        allTargetNames.Add(hostName)
                        For Each childName In childNamesForInstance
                            setChildren.Add(childName)
                            allTargetNames.Add(childName)
                            childNames.Add(childName)
                        Next
                    Next

                Catch ex As Exception
                    If Not IsNoTxnNoise(ex.Message) Then
                        scanFails.Add(f.Name)
                    End If
                Finally
                    If hostDoc IsNot Nothing Then
                        Try : hostDoc.Close(False) : Catch : End Try
                    End If
                End Try
            Next

            ' 복합 패밀리(상위) 목록 (리포트용)
            Dim compositeFamilies As New List(Of Family)()
            For Each pname In parentToChildren.Keys
                Dim fam As Family = Nothing
                If nameToFamily.TryGetValue(pname, fam) Then
                    compositeFamilies.Add(fam)
                End If
            Next

            Dim standaloneFamilyNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If includeStandaloneFamilies Then
                For Each f In allEditable
                    If f Is Nothing OrElse f.FamilyCategory Is Nothing Then Continue For
                    If IsAnnotationFamily(f) Then Continue For

                    Dim famName As String = If(f.Name, String.Empty)
                    If String.IsNullOrWhiteSpace(famName) Then Continue For
                    If allTargetNames.Contains(famName) Then Continue For

                    standaloneFamilyNames.Add(famName)
                    allTargetNames.Add(famName)
                Next
            End If

            '----- 2. 그래프를 하위 → 상위 순으로 정렬 (DFS, 이름 기준) -----
            Dim order As New List(Of String)()
            Dim mark As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase) ' 0:미방문,1:방문중,2:완료

            Dim dfs As Action(Of String) =
                Sub(name As String)
                    Dim st As Integer = 0
                    If mark.TryGetValue(name, st) Then
                        If st = 2 Then Return
                        If st = 1 Then
                            ' 순환 구조가 있다면 더 들어가지 않고 끊음
                            Return
                        End If
                    End If

                    mark(name) = 1
                    Dim childs As HashSet(Of String) = Nothing
                    If parentToChildren.TryGetValue(name, childs) Then
                        For Each cn In childs
                            dfs(cn)
                        Next
                    End If
                    mark(name) = 2
                    order.Add(name)   ' 자식 먼저 처리 후 자기 자신 ⇒ 하위 → 상위
                End Sub

            For Each name In allTargetNames
                dfs(name)
            Next

            '----- 3. 계층 역순으로 파라미터 추가 & 연동 -----
            Dim fatalEx As Exception = Nothing
            Dim addedChild As Integer = 0
            Dim addedHost As Integer = 0
            Dim addedStandalone As Integer = 0
            Dim linkCnt As Integer = 0
            Dim verifyOk As Integer = 0
            Dim verifyFail As Integer = 0
            Dim skipTotal As Integer = 0
            Dim parentFails As New List(Of String)()
            Dim childFails As New List(Of String)()
            Dim standaloneFails As New List(Of String)()
            Dim skips As New List(Of String)()
            Dim compositeSuccessCount As Integer = 0

            Dim applyIdx As Integer = 0
            Dim applyTotal As Integer = order.Count
            report("apply", 0.0R, 0, applyTotal, "파라미터 적용 준비", "")

            Using tgAll As New TransactionGroup(doc, "KKY Shared Param Propagate")
                tgAll.Start()
                Try
                    ' order: 하위 → 상위. leaf(A-6) 먼저 처리.
                    For Each famName In order
                        applyIdx += 1
                        report("apply", If(applyTotal <= 0, 0.0R, CDbl(applyIdx) / CDbl(applyTotal)), applyIdx, applyTotal, "적용 중...", famName)
                        Dim fam As Family = Nothing
                        If Not nameToFamily.TryGetValue(famName, fam) Then Continue For

                        Dim isParent As Boolean = parentToChildren.ContainsKey(famName)
                        Dim isChild As Boolean = childNames.Contains(famName)
                        Dim isStandalone As Boolean = standaloneFamilyNames.Contains(famName)

                        ProcessFamilyBottomUp(doc,
                                             fam,
                                             extDefs,
                                             paramNames,
                                             parentToChildren,
                                             excludeDummy,
                                             chosenIsInstance,
                                             chosenPG,
                                             isParent,
                                             isChild,
                                             isStandalone,
                                             addedHost,
                                             addedChild,
                                             addedStandalone,
                                              linkCnt,
                                              verifyOk,
                                              verifyFail,
                                              skipTotal,
                                              parentFails,
                                              childFails,
                                              standaloneFails,
                                              skips,
                                              compositeSuccessCount)
                    Next

                    tgAll.Assimilate()
                Catch ex As Exception
                    fatalEx = ex
                    Try : tgAll.RollBack() : Catch : End Try
                End Try
            End Using

            If fatalEx IsNot Nothing Then
                result.Message = fatalEx.Message
                Return RunStatus.Failed
            End If

            '----- 4. 리포트 구성 -----
            Dim header As New StringBuilder()
            header.AppendLine($"패밀리 스캔: {totalEditableCount}개 / 복합 패밀리: {compositeFamilies.Count}개 / 성공: {compositeSuccessCount}개")
            header.AppendLine($"하위 패밀리 파라미터 추가: {addedChild}")
            header.AppendLine($"복합 패밀리 파라미터 추가/교정: {addedHost}")
            header.AppendLine($"단일 패밀리 파라미터 추가: {addedStandalone} / 대상: {standaloneFamilyNames.Count}")
            header.AppendLine($"파라미터 연동 성공 카운트: {linkCnt}")
            header.AppendLine($"연동 검증 OK: {verifyOk} / 미연동: {verifyFail}")
            header.AppendLine($"스킵(연동 불가) 건수: {skipTotal}")
            header.AppendLine($"선택 모드: {(If(chosenIsInstance, "인스턴스", "타입"))}")
            header.AppendLine($"단일 패밀리 추가 옵션: {(If(includeStandaloneFamilies, "사용", "미사용"))}")
            header.AppendLine($"Dummy 제외된 하위 패밀리 수: {dummyExcludedCount}")

            Dim skipLines = skips.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
            Dim failLines = parentFails.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
            failLines.AddRange(standaloneFails.Distinct(StringComparer.OrdinalIgnoreCase))
            Dim scanLines = scanFails.Distinct(StringComparer.OrdinalIgnoreCase).ToList()

            Dim hasDetail As Boolean = (scanLines.Count + skipLines.Count + failLines.Count + childFails.Count > 0)
            If hasDetail Then
                header.AppendLine().AppendLine("상세:")
                If scanLines.Count > 0 Then
                    header.AppendLine("스캔 실패(패밀리 편집 불가) 목록:")
                    header.AppendLine(String.Join(vbCrLf, scanLines.Take(80)))
                    If scanLines.Count > 80 Then header.AppendLine("...")
                End If
                If skipLines.Count > 0 Then
                    header.AppendLine("스킵(연동 불가) 목록:")
                    header.AppendLine(String.Join(vbCrLf, skipLines.Take(80)))
                    If skipLines.Count > 80 Then header.AppendLine("...")
                End If
                If failLines.Count + childFails.Count > 0 Then
                    header.AppendLine("오류 목록:")
                    Dim failMerged As New List(Of String)()
                    failMerged.AddRange(failLines)
                    failMerged.AddRange(childFails)
                    header.AppendLine(String.Join(vbCrLf, failMerged.Take(80)))
                    If failMerged.Count > 80 Then header.AppendLine("...")
                End If
            End If

            result.Report = header.ToString()
            result.Details = BuildDetails(scanLines, skipLines, failLines, childFails)
            Return RunStatus.Succeeded
        End Function

        ' 패밀리 1개 단위 처리 (계층 역순에서 호출)
        Private Shared Sub ProcessFamilyBottomUp(projDoc As Document,
                                                 fam As Family,
                                                 extDefs As IEnumerable(Of ExternalDefinition),
                                                 paramNames As List(Of String),
                                                 parentToChildren As Dictionary(Of String, HashSet(Of String)),
                                                 excludeDummy As Boolean,
                                                 chosenIsInstance As Boolean,
                                                 chosenPG As BuiltInParameterGroup,
                                                 isParent As Boolean,
                                                 isChild As Boolean,
                                                 isStandalone As Boolean,
                                                 ByRef addedHost As Integer,
                                                 ByRef addedChild As Integer,
                                                 ByRef addedStandalone As Integer,
                                                 ByRef linkCnt As Integer,
                                                 ByRef verifyOk As Integer,
                                                 ByRef verifyFail As Integer,
                                                 ByRef skipTotal As Integer,
                                                 parentFails As List(Of String),
                                                 childFails As List(Of String),
                                                 standaloneFails As List(Of String),
                                                 skips As List(Of String),
                                                 ByRef compositeSuccessCount As Integer)

            If fam Is Nothing Then Return
            If IsAnnotationFamily(fam) Then Return

            Dim famName As String = fam.Name
            Dim hasChildren As Boolean = parentToChildren.ContainsKey(famName)

            Dim famDoc As Document = Nothing
            Dim ok As Boolean = True
            Dim localAdded As Integer = 0
            Dim localSkipAssoc As Integer = 0

            Try
                famDoc = projDoc.EditFamily(fam)
                Dim fm As FamilyManager = famDoc.FamilyManager

                ' 1) 이 패밀리에 공유 파라미터 추가/교정
                For Each ed In extDefs
                    Dim res = EnsureSharedParamInFamily(famDoc, ed, chosenIsInstance, chosenPG)
                    If res.Added Then localAdded += 1
                    If Not String.IsNullOrEmpty(res.ErrorMessage) Then
                        ok = False
                        If isParent Then
                            parentFails.Add(famName & ": " & res.ErrorMessage)
                        ElseIf isChild Then
                            childFails.Add(famName & ": " & res.ErrorMessage)
                        ElseIf isStandalone Then
                            standaloneFails.Add(famName & ": " & res.ErrorMessage)
                        End If
                    End If
                    Try : famDoc.Regenerate() : Catch : End Try
                Next

                If localAdded > 0 Then
                    If isChild Then addedChild += localAdded
                    If isParent Then addedHost += localAdded
                    If isStandalone Then addedStandalone += localAdded
                End If

                ' 2) 자식이 있는 패밀리라면 하위 인스턴스와 연동
                If hasChildren Then
                    Dim childrenOfHost As HashSet(Of String) = Nothing
                    parentToChildren.TryGetValue(famName, childrenOfHost)

                    Dim hostParams As New Dictionary(Of String, FamilyParameter)(StringComparer.OrdinalIgnoreCase)
                    For Each p As FamilyParameter In fm.Parameters
                        If p.Definition IsNot Nothing Then
                            Dim nm = p.Definition.Name
                            If paramNames.Contains(nm, StringComparer.OrdinalIgnoreCase) Then
                                hostParams(nm) = p
                            End If
                        End If
                    Next

                    Using t As New Transaction(famDoc, "KKY: Associate nested shared params")
                        t.Start()

                        Dim insts = New FilteredElementCollector(famDoc).
                                    OfClass(GetType(FamilyInstance)).
                                    Cast(Of FamilyInstance)().
                                    ToList()

                        For Each fi In insts
                            Dim childF As Family = Nothing
                            Try
                                Dim sym As FamilySymbol = TryCast(famDoc.GetElement(fi.Symbol.Id), FamilySymbol)
                                If sym Is Nothing Then Continue For
                                childF = sym.Family
                            Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                                ok = False
                                parentFails.Add($"{famName}: [Associate] 오류 - {ex.Message}")
                                Continue For
                            Catch ex As Exception
                                ok = False
                                parentFails.Add($"{famName}: [Associate] 오류 - {ex.Message}")
                                Continue For
                            End Try

                            If childF Is Nothing Then Continue For
                            If IsAnnotationFamily(childF) Then Continue For
                            If ShouldSkipDummy(excludeDummy, childF) Then Continue For

                            ' 그래프상 이 호스트의 자식으로 등록된 패밀리만 대상
                            If childrenOfHost IsNot Nothing AndAlso childrenOfHost.Count > 0 AndAlso
                               Not childrenOfHost.Contains(childF.Name) Then
                                Continue For
                            End If

                            For Each name In paramNames
                                Try
                                    Dim hostParam As FamilyParameter = Nothing
                                    If Not hostParams.TryGetValue(name, hostParam) OrElse hostParam Is Nothing Then
                                        Continue For
                                    End If

                                    Dim p As Parameter = TryGetElementParameterByName(fi, name)
                                    If p Is Nothing Then
                                        TraceAssociateSkipDebug(famDoc, famName, childF, fi, name, hostParam, "child-param-missing", p, fm)
                                        localSkipAssoc += 1
                                        Continue For
                                    End If
                                    If p.IsReadOnly Then
                                        TraceAssociateSkipDebug(famDoc, famName, childF, fi, name, hostParam, "child-param-readonly", p, fm)
                                        localSkipAssoc += 1
                                        Continue For
                                    End If
                                    If Not famDoc.FamilyManager.CanElementParameterBeAssociated(p) Then
                                        TraceAssociateSkipDebug(famDoc, famName, childF, fi, name, hostParam, "child-param-not-associable", p, fm)
                                        localSkipAssoc += 1
                                        Continue For
                                    End If

                                    Try
                                        famDoc.FamilyManager.AssociateElementParameterToFamilyParameter(p, hostParam)
                                        linkCnt += 1
                                    Catch
                                        ok = False
                                        parentFails.Add($"{famName}: [Associate] 실패 - {name}")
                                    End Try
                                Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                                    ok = False
                                    parentFails.Add($"{famName}: [Associate] 오류 - {ex.Message}")
                                    Continue For
                                Catch ex As Exception
                                    ok = False
                                    parentFails.Add($"{famName}: [Associate] 오류 - {ex.Message}")
                                    Continue For
                                End Try
                            Next
                        Next

                        t.Commit()
                    End Using

                    Try : famDoc.Regenerate() : Catch : End Try

                    ' 3) 연동 검증 (그래프상의 자식만 대상으로)
                    Dim v = VerifyAssociations(famDoc, hostParams, paramNames, excludeDummy, childrenOfHost)
                    verifyOk += v.Ok : verifyFail += v.Fail

                    Dim hasAnySuccess As Boolean = (v.Ok > 0)
                    Dim localSkipTotal As Integer = localSkipAssoc + v.Skip

                    If Not hasAnySuccess Then
                        skipTotal += localSkipTotal
                    End If

                    ' 패밀리 안에서 한 번이라도 연동이 성공했다면
                    ' 상위 "Error / Skip 목록"에는 올리지 않는다 (부분 실패는 숫자 카운트로만 유지)
                    If v.FailItems IsNot Nothing AndAlso v.FailItems.Count > 0 AndAlso Not hasAnySuccess Then
                        ok = False
                        parentFails.Add(famName)
                    End If
                    If v.SkipItems IsNot Nothing AndAlso v.SkipItems.Count > 0 AndAlso Not hasAnySuccess Then
                        skips.Add(famName)
                    End If
                End If

                ' 4) 프로젝트에 로드
                TxnUtil.SafeLoadFamily(
                    famDoc,
                    projDoc,
                    $"KKY Load '{famName}'",
                    _saveModifiedFamilies,
                    _modifiedFamilyOutputFolder)
                If ok AndAlso hasChildren Then
                    compositeSuccessCount += 1
                End If

            Catch ex As Exception
                ok = False
                If Not IsNoTxnNoise(ex.Message) Then
                    If isParent Then
                        parentFails.Add(famName)
                    ElseIf isChild Then
                        childFails.Add(famName)
                    ElseIf isStandalone Then
                        standaloneFails.Add(famName)
                    End If
                End If
            Finally
                If famDoc IsNot Nothing Then
                    Try : famDoc.Close(False) : Catch : End Try
                End If
            End Try
        End Sub

        '==================== 공통 헬퍼들 ====================

        Private Shared Sub SaveChildFamilyImmediatelyAfterAdd(famDoc As Document,
                                                              famName As String,
                                                              extDefs As IEnumerable(Of ExternalDefinition),
                                                              chosenPG As BuiltInParameterGroup)
            If famDoc Is Nothing Then Return

            Dim fm As FamilyManager = Nothing
            Try : fm = famDoc.FamilyManager : Catch : fm = Nothing : End Try

            If extDefs IsNot Nothing Then
                For Each extDef In extDefs
                    If extDef Is Nothing Then Continue For
                    Dim param As FamilyParameter = FindFamilyParameterByNameForSaveDebug(fm, extDef.Name)
                    TraceChildSaveDebug(famDoc, "before-save", famName, extDef, chosenPG, param)
                Next
            End If

            Dim saveMode As String = "none"
            Dim savePath As String = ""
            Dim saveFailureMessage As String = ""

            Try
                Dim docPath As String = ""
                Try : docPath = If(famDoc.PathName, String.Empty) : Catch : docPath = "" : End Try

                If Not String.IsNullOrWhiteSpace(docPath) Then
                    Try
                        famDoc.Save()
                        saveMode = "save"
                        savePath = docPath
                    Catch ex As Exception
                        saveFailureMessage = ex.Message
                    End Try
                End If

                If saveMode = "none" Then
                    Dim tempDir As String = Path.Combine(Path.GetTempPath(), "KKY_Tool_Revit", "ParamPropChildSave")
                    Directory.CreateDirectory(tempDir)

                    Dim baseName As String = Path.GetFileNameWithoutExtension(If(famName, famDoc.Title))
                    If String.IsNullOrWhiteSpace(baseName) Then baseName = "ChildFamily"

                    For Each ch As Char In Path.GetInvalidFileNameChars()
                        baseName = baseName.Replace(ch, "_"c)
                    Next

                    savePath = Path.Combine(tempDir, $"{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}.rfa")
                    Dim options As New SaveAsOptions()
                    options.OverwriteExistingFile = True
                    famDoc.SaveAs(savePath, options)
                    saveMode = If(String.IsNullOrWhiteSpace(saveFailureMessage), "saveas", "saveas-fallback")
                End If

                UI.Hub.UiBridgeExternalEvent.HostLog(
                    "debug",
                    $"[paramprop][child-save-debug] stage=saved; family={famName}; mode={saveMode}; path=""{savePath}""; saveMessage={If(String.IsNullOrWhiteSpace(saveFailureMessage), "<none>", saveFailureMessage)}")
            Catch ex As Exception
                UI.Hub.UiBridgeExternalEvent.HostLog(
                    "debug",
                    $"[paramprop][child-save-debug] stage=save-fail; family={famName}; message={ex.Message}")
                Return
            End Try

            Try : famDoc.Regenerate() : Catch : End Try

            If extDefs IsNot Nothing Then
                Try : fm = famDoc.FamilyManager : Catch : fm = Nothing : End Try

                For Each extDef In extDefs
                    If extDef Is Nothing Then Continue For
                    Dim param As FamilyParameter = FindFamilyParameterByNameForSaveDebug(fm, extDef.Name)
                    TraceChildSaveDebug(famDoc, "after-save", famName, extDef, chosenPG, param)
                Next
            End If
        End Sub

        Private Shared Sub SaveParentFamilyImmediatelyAfterProjectLoad(famDoc As Document,
                                                                       famName As String,
                                                                       extDefs As IEnumerable(Of ExternalDefinition),
                                                                       chosenPG As BuiltInParameterGroup)
            If famDoc Is Nothing Then Return

            Dim fm As FamilyManager = Nothing
            Try : fm = famDoc.FamilyManager : Catch : fm = Nothing : End Try

            If extDefs IsNot Nothing Then
                For Each extDef In extDefs
                    If extDef Is Nothing Then Continue For
                    Dim param As FamilyParameter = FindFamilyParameterByNameForSaveDebug(fm, extDef.Name)
                    TraceParentSaveDebug(famDoc, "before-save", famName, extDef, chosenPG, param)
                Next
            End If

            Dim saveMode As String = "none"
            Dim savePath As String = ""
            Dim saveFailureMessage As String = ""

            Try
                Dim docPath As String = ""
                Try : docPath = If(famDoc.PathName, String.Empty) : Catch : docPath = "" : End Try

                If _saveModifiedFamilies AndAlso Not String.IsNullOrWhiteSpace(_modifiedFamilyOutputFolder) Then
                    Directory.CreateDirectory(_modifiedFamilyOutputFolder)

                    Dim baseName As String = Path.GetFileNameWithoutExtension(If(famName, famDoc.Title))
                    If String.IsNullOrWhiteSpace(baseName) Then baseName = "ParentFamily"

                    For Each ch As Char In Path.GetInvalidFileNameChars()
                        baseName = baseName.Replace(ch, "_"c)
                    Next

                    savePath = Path.Combine(_modifiedFamilyOutputFolder, $"{baseName}.rfa")
                    Dim options As New SaveAsOptions()
                    options.OverwriteExistingFile = True
                    famDoc.SaveAs(savePath, options)
                    saveMode = "saveas-output"
                ElseIf Not String.IsNullOrWhiteSpace(docPath) Then
                    Try
                        famDoc.Save()
                        saveMode = "save"
                        savePath = docPath
                    Catch ex As Exception
                        saveFailureMessage = ex.Message
                    End Try
                End If

                If saveMode = "none" Then
                    Dim tempDir As String = Path.Combine(Path.GetTempPath(), "KKY_Tool_Revit", "ParamPropParentSave")
                    Directory.CreateDirectory(tempDir)

                    Dim baseName As String = Path.GetFileNameWithoutExtension(If(famName, famDoc.Title))
                    If String.IsNullOrWhiteSpace(baseName) Then baseName = "ParentFamily"

                    For Each ch As Char In Path.GetInvalidFileNameChars()
                        baseName = baseName.Replace(ch, "_"c)
                    Next

                    savePath = Path.Combine(tempDir, $"{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}.rfa")
                    Dim options As New SaveAsOptions()
                    options.OverwriteExistingFile = True
                    famDoc.SaveAs(savePath, options)
                    saveMode = If(String.IsNullOrWhiteSpace(saveFailureMessage), "saveas", "saveas-fallback")
                End If

                If String.IsNullOrWhiteSpace(savePath) OrElse Not File.Exists(savePath) Then
                    Throw New IOException($"저장 후 파일을 찾을 수 없습니다: {savePath}")
                End If

                UI.Hub.UiBridgeExternalEvent.HostLog(
                    "debug",
                    $"[paramprop][parent-save-debug] stage=saved; family={famName}; mode={saveMode}; path=""{savePath}""; saveMessage={If(String.IsNullOrWhiteSpace(saveFailureMessage), "<none>", saveFailureMessage)}")
            Catch ex As Exception
                UI.Hub.UiBridgeExternalEvent.HostLog(
                    "debug",
                    $"[paramprop][parent-save-debug] stage=save-fail; family={famName}; message={ex.Message}")
                Return
            End Try

            Try : famDoc.Regenerate() : Catch : End Try

            If extDefs IsNot Nothing Then
                Try : fm = famDoc.FamilyManager : Catch : fm = Nothing : End Try

                For Each extDef In extDefs
                    If extDef Is Nothing Then Continue For
                    Dim param As FamilyParameter = FindFamilyParameterByNameForSaveDebug(fm, extDef.Name)
                    TraceParentSaveDebug(famDoc, "after-save", famName, extDef, chosenPG, param)
                Next
            End If
        End Sub

        Private Shared Function FindFamilyParameterByNameForSaveDebug(fm As FamilyManager,
                                                                      name As String) As FamilyParameter
            If fm Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then Return Nothing

            Try
                Return fm.Parameters.Cast(Of FamilyParameter)().
                    FirstOrDefault(Function(x) x IsNot Nothing AndAlso
                                              x.Definition IsNot Nothing AndAlso
                                              String.Equals(x.Definition.Name, name, StringComparison.OrdinalIgnoreCase))
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub TraceChildSaveDebug(famDoc As Document,
                                               stage As String,
                                               famName As String,
                                               extDef As ExternalDefinition,
                                               chosenPG As BuiltInParameterGroup,
                                               param As FamilyParameter)
            Try
#If REVIT2025 Then
                Dim fm As FamilyManager = Nothing
                Try : fm = famDoc.FamilyManager : Catch : fm = Nothing : End Try
                Dim requestedGroupTypeId As ForgeTypeId = ResolveUserAssignableGroupTypeId(fm, chosenPG, extDef)
                Dim actualGroupTypeId As ForgeTypeId = TryGetFamilyParameterGroupTypeIdForSaveDebug(param)
                UI.Hub.UiBridgeExternalEvent.HostLog(
                    "debug",
                    $"[paramprop][child-save-debug] stage={stage}; family={famName}; param={extDef.Name}; requested={DescribeGroupTypeIdForSaveDebug(requestedGroupTypeId)}; actual={DescribeGroupTypeIdForSaveDebug(actualGroupTypeId)}; paramFound={param IsNot Nothing}")
#Else
                Dim actualGroupName As String = "<null>"
                Try
                    If param IsNot Nothing AndAlso param.Definition IsNot Nothing Then
                        actualGroupName = param.Definition.ParameterGroup.ToString()
                    End If
                Catch
                End Try
                UI.Hub.UiBridgeExternalEvent.HostLog(
                    "debug",
                    $"[paramprop][child-save-debug] stage={stage}; family={famName}; param={extDef.Name}; requested={chosenPG}; actual={actualGroupName}; paramFound={param IsNot Nothing}")
#End If
            Catch
            End Try
        End Sub

        Private Shared Sub TraceParentSaveDebug(famDoc As Document,
                                                stage As String,
                                                famName As String,
                                                extDef As ExternalDefinition,
                                                chosenPG As BuiltInParameterGroup,
                                                param As FamilyParameter)
            Try
#If REVIT2025 Then
                Dim fm As FamilyManager = Nothing
                Try : fm = famDoc.FamilyManager : Catch : fm = Nothing : End Try
                Dim requestedGroupTypeId As ForgeTypeId = ResolveUserAssignableGroupTypeId(fm, chosenPG, extDef)
                Dim actualGroupTypeId As ForgeTypeId = TryGetFamilyParameterGroupTypeIdForSaveDebug(param)
                UI.Hub.UiBridgeExternalEvent.HostLog(
                    "debug",
                    $"[paramprop][parent-save-debug] stage={stage}; family={famName}; param={extDef.Name}; requested={DescribeGroupTypeIdForSaveDebug(requestedGroupTypeId)}; actual={DescribeGroupTypeIdForSaveDebug(actualGroupTypeId)}; paramFound={param IsNot Nothing}")
#Else
                Dim actualGroupName As String = "<null>"
                Try
                    If param IsNot Nothing AndAlso param.Definition IsNot Nothing Then
                        actualGroupName = param.Definition.ParameterGroup.ToString()
                    End If
                Catch
                End Try
                UI.Hub.UiBridgeExternalEvent.HostLog(
                    "debug",
                    $"[paramprop][parent-save-debug] stage={stage}; family={famName}; param={extDef.Name}; requested={chosenPG}; actual={actualGroupName}; paramFound={param IsNot Nothing}")
#End If
            Catch
            End Try
        End Sub

#If REVIT2025 Then
        Private Shared Function TryGetFamilyParameterGroupTypeIdForSaveDebug(param As FamilyParameter) As ForgeTypeId
            If param Is Nothing OrElse param.Definition Is Nothing Then Return Nothing

            Try
                Return param.Definition.GetGroupTypeId()
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function DescribeGroupTypeIdForSaveDebug(groupTypeId As ForgeTypeId) As String
            If groupTypeId Is Nothing Then Return "<null>"

            Try
                Dim raw As String = groupTypeId.TypeId
                If String.IsNullOrWhiteSpace(raw) Then Return "<null>"
                Return raw
            Catch
                Return "<null>"
            End Try
        End Function
#End If

#If REVIT2025 Then
        '==================== Revit 2025: Parameter groupTypeId mapping (FIX) ====================
        ' 기존 인덱스 기반 매핑(CInt(enum) → list index) 제거
        Private Shared Function ResolveUserAssignableGroupTypeId(
            fm As FamilyManager,
            groupPG As BuiltInParameterGroup,
            extDef As ExternalDefinition
        ) As ForgeTypeId

            Dim gtid As ForgeTypeId = Nothing

            ' 1) compat layer를 통해 직접 GroupTypeId로 매핑 (동적 1000+ 그룹 id 포함)
            Try
                gtid = BuiltInParameterGroupCompat.ToGroupTypeId(groupPG)
            Catch
                gtid = Nothing
            End Try

            ' 2) fallback: 정의 자체의 그룹
            If gtid Is Nothing Then
                Try : gtid = extDef.GetGroupTypeId() : Catch : gtid = Nothing : End Try
            End If

            ' 3) 최후 fallback
            If gtid Is Nothing Then
                Try : gtid = GroupTypeId.Data : Catch : End Try
            End If
            If gtid Is Nothing Then gtid = New ForgeTypeId()

            Return gtid
        End Function
#End If

        Private Structure EnsureResult
            Public Added As Boolean
            Public FinalOk As Boolean
            Public HadMismatch As Boolean
            Public ErrorMessage As String
        End Structure

        Private Shared Function EnsureSharedParamInFamily(famDoc As Document,
                                                         extDef As ExternalDefinition,
                                                         isInstance As Boolean,
                                                         groupPG As BuiltInParameterGroup) As EnsureResult
            Dim fm As FamilyManager = famDoc.FamilyManager
            Dim result As New EnsureResult With {
                .Added = False,
                .FinalOk = False,
                .HadMismatch = False,
                .ErrorMessage = Nothing
            }

            Dim anyByName As FamilyParameter =
                fm.Parameters.Cast(Of FamilyParameter)().
                FirstOrDefault(Function(p) p.Definition IsNot Nothing AndAlso
                                           String.Equals(p.Definition.Name, extDef.Name, StringComparison.OrdinalIgnoreCase))

            If anyByName Is Nothing Then
                Try
                    TxnUtil.WithTxn(famDoc, $"TEMP non-shared add: {extDef.Name}",
                        Sub()
#If REVIT2019 Or REVIT2021 Or REVIT2023 Then
                            ' shared parameter directly
                            fm.AddParameter(extDef, groupPG, isInstance)
#ElseIf REVIT2025 Then
                            Dim gtid As ForgeTypeId = ResolveUserAssignableGroupTypeId(fm, groupPG, extDef)
                            fm.AddParameter(extDef, gtid, isInstance)
#End If
                        End Sub)
                    result.Added = True
                    Try : famDoc.Regenerate() : Catch : End Try
                    anyByName =
                        fm.Parameters.Cast(Of FamilyParameter)().
                        FirstOrDefault(Function(x) x.Definition IsNot Nothing AndAlso
                                                   String.Equals(x.Definition.Name, extDef.Name, StringComparison.OrdinalIgnoreCase))
                Catch ex As Exception
                    result.ErrorMessage = $"임시 비공유 추가 실패: {ex.Message}"
                    Return result
                End Try
            End If

            If anyByName Is Nothing Then
                result.ErrorMessage = "동일 이름 파라미터를 찾을 수 없음(임시 생성 실패)."


                Return result
            End If

            If IsExactSharedFamilyParameter(anyByName, extDef, isInstance, groupPG, fm) Then
                result.FinalOk = True
                Return result
            End If

            Dim replaceFailed As Boolean = False
            Try
                TxnUtil.WithTxn(famDoc, $"Replace to shared: {extDef.Name}",
                    Sub()
#If REVIT2019 Or REVIT2021 Or REVIT2023 Then
                        fm.ReplaceParameter(anyByName, extDef, groupPG, isInstance)
#ElseIf REVIT2025 Then
                        Dim gtid As ForgeTypeId = ResolveUserAssignableGroupTypeId(fm, groupPG, extDef)
                        fm.ReplaceParameter(anyByName, extDef, gtid, isInstance)
#End If
                    End Sub)
            Catch
                replaceFailed = True
            End Try

            Dim corrected As FamilyParameter =
                fm.Parameters.Cast(Of FamilyParameter)().
                FirstOrDefault(Function(x) x.Definition IsNot Nothing AndAlso
                                           String.Equals(x.Definition.Name, extDef.Name, StringComparison.OrdinalIgnoreCase))
#If REVIT2019 Or REVIT2021 Or REVIT2023 Then
            Dim okNow As Boolean =
                (corrected IsNot Nothing AndAlso corrected.IsShared AndAlso corrected.IsInstance = isInstance AndAlso
                 corrected.Definition.ParameterGroup = groupPG)
#ElseIf REVIT2025 Then
            Dim groupOk As Boolean = True
            Try
                groupOk = (corrected IsNot Nothing AndAlso corrected.Definition IsNot Nothing AndAlso
                           corrected.Definition.GetGroupTypeId().Equals(ResolveUserAssignableGroupTypeId(fm, groupPG, extDef)))
            Catch
                groupOk = True
            End Try

            Dim okNow As Boolean =
                (corrected IsNot Nothing AndAlso corrected.IsShared AndAlso corrected.IsInstance = isInstance AndAlso groupOk)
#End If

            If okNow AndAlso Not replaceFailed Then
                result.FinalOk = True
                Return result
            End If

            result.HadMismatch = True
            Try
                TxnUtil.WithTxn(famDoc, $"HardFix Remove/Add: {extDef.Name}",
                    Sub()
                        Try
                            fm.RemoveParameter(anyByName)
                        Catch remEx As Exception
                            Throw New InvalidOperationException(
                                $"Remove 실패(레이블/치수/공식 참조 중일 수 있음): {remEx.Message}")
                        End Try
#If REVIT2019 Or REVIT2021 Or REVIT2023 Then
                        fm.AddParameter(extDef, groupPG, isInstance)
#ElseIf REVIT2025 Then
                            Dim gtid As ForgeTypeId = ResolveUserAssignableGroupTypeId(fm, groupPG, extDef)
                            fm.AddParameter(extDef, gtid, isInstance)
#End If
                    End Sub)
                Try : famDoc.Regenerate() : Catch : End Try
            Catch ex As Exception
                result.ErrorMessage = $"HardFix 실패: {ex.Message}"
                Return result
            End Try

            corrected =
                fm.Parameters.Cast(Of FamilyParameter)().
                FirstOrDefault(Function(x) x.Definition IsNot Nothing AndAlso
                                           x.IsShared AndAlso
                                           String.Equals(x.Definition.Name, extDef.Name, StringComparison.OrdinalIgnoreCase))
#If REVIT2019 Or REVIT2021 Or REVIT2023 Then
            result.FinalOk = (corrected IsNot Nothing AndAlso corrected.Definition.ParameterGroup = groupPG)
#ElseIf REVIT2025 Then
            Dim groupOk2 As Boolean = True
            Try
                groupOk2 = (corrected IsNot Nothing AndAlso corrected.Definition IsNot Nothing AndAlso
                            corrected.Definition.GetGroupTypeId().Equals(ResolveUserAssignableGroupTypeId(fm, groupPG, extDef)))
            Catch
                groupOk2 = True
            End Try
            result.FinalOk = (corrected IsNot Nothing AndAlso groupOk2)
#End If
            Return result
        End Function

        Private Shared Function IsExactSharedFamilyParameter(fp As FamilyParameter,
                                                            extDef As ExternalDefinition,
                                                            isInstance As Boolean,
                                                            groupPG As BuiltInParameterGroup,
                                                            fm As FamilyManager) As Boolean
            If fp Is Nothing OrElse extDef Is Nothing OrElse fp.Definition Is Nothing Then Return False
            If Not String.Equals(fp.Definition.Name, extDef.Name, StringComparison.OrdinalIgnoreCase) Then Return False

            Dim isShared As Boolean = False
            Try : isShared = fp.IsShared : Catch : isShared = False : End Try
            If Not isShared Then Return False

            Dim scopeOk As Boolean = False
            Try : scopeOk = (fp.IsInstance = isInstance) : Catch : scopeOk = False : End Try
            If Not scopeOk Then Return False

            Dim existingGuid As Guid = Guid.Empty
            If Not TryGetFamilyParameterGuid(fp, existingGuid) Then Return False
            If existingGuid <> extDef.GUID Then Return False

#If REVIT2019 Or REVIT2021 Or REVIT2023 Then
            Try
                Return fp.Definition.ParameterGroup = groupPG
            Catch
                Return False
            End Try
#ElseIf REVIT2025 Then
            Try
                Dim requestedGroupTypeId As ForgeTypeId = ResolveUserAssignableGroupTypeId(fm, groupPG, extDef)
                Return fp.Definition.GetGroupTypeId().Equals(requestedGroupTypeId)
            Catch
                Return True
            End Try
#End If
        End Function

        Private Shared Function TryGetFamilyParameterGuid(fp As FamilyParameter, ByRef guid As Guid) As Boolean
            guid = Guid.Empty
            If fp Is Nothing Then Return False

            Try
                Dim prop = fp.GetType().GetProperty("GUID",
                                                    System.Reflection.BindingFlags.Public Or
                                                    System.Reflection.BindingFlags.Instance)
                If prop Is Nothing Then Return False

                Dim value = prop.GetValue(fp, Nothing)
                If TypeOf value Is Guid Then
                    guid = DirectCast(value, Guid)
                    Return guid <> Guid.Empty
                End If
            Catch
            End Try

            Return False
        End Function

        Private Structure VerifyResult
            Public Ok As Integer
            Public Fail As Integer
            Public FailItems As List(Of String)
            Public Skip As Integer
            Public SkipItems As List(Of String)
        End Structure

        Private Structure TypeDrivenAssociationResult
            Public Applicable As Boolean
            Public Success As Boolean
        End Structure

        Private Shared Function VerifyAssociations(hostDoc As Document,
                                                   hostParams As Dictionary(Of String, FamilyParameter),
                                                   paramNames As List(Of String),
                                                   excludeDummy As Boolean,
                                                   allowedChildNames As HashSet(Of String)) As VerifyResult

            Dim fm = hostDoc.FamilyManager
            Dim okCnt As Integer = 0
            Dim failCnt As Integer = 0
            Dim skipCnt As Integer = 0
            Dim fails As New List(Of String)()
            Dim skipItems As New List(Of String)()

            Dim insts = New FilteredElementCollector(hostDoc).
                        OfClass(GetType(FamilyInstance)).
                        Cast(Of FamilyInstance)().
                        ToList()

            For Each fi In insts
                Dim childF As Family = Nothing
                Try
                    Dim sym As FamilySymbol = TryCast(hostDoc.GetElement(fi.Symbol.Id), FamilySymbol)
                    If sym Is Nothing Then Continue For
                    childF = sym.Family
                Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                    failCnt += 1
                    fails.Add($"{fi.Id.IntValue()}:(verify-exception: {ex.Message})")
                    Continue For
                Catch ex As Exception
                    failCnt += 1
                    fails.Add($"{fi.Id.IntValue()}:(verify-exception: {ex.Message})")
                    Continue For
                End Try

                If childF Is Nothing Then Continue For
                If IsAnnotationFamily(childF) Then Continue For
                If ShouldSkipDummy(excludeDummy, childF) Then Continue For

                ' 그래프 상 이 호스트의 자식으로 등록된 패밀리만 검증
                If allowedChildNames IsNot Nothing AndAlso allowedChildNames.Count > 0 AndAlso
                   Not allowedChildNames.Contains(childF.Name) Then
                    Continue For
                End If

                For Each name In paramNames
                    Try
                        Dim hostParam As FamilyParameter = Nothing
                        If Not hostParams.TryGetValue(name, hostParam) OrElse hostParam Is Nothing Then
                            failCnt += 1
                            fails.Add($"{fi.Id.IntValue()}:{name} (host-missing)")
                            Continue For
                        End If

                        Dim p As Parameter = Nothing
                        Try
                            p = TryGetElementParameterByName(fi, name)
                        Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                            failCnt += 1
                            fails.Add($"{fi.Id.IntValue()}:{name} (verify-exception: {ex.Message})")
                            Continue For
                        Catch ex As Exception
                            failCnt += 1
                            fails.Add($"{fi.Id.IntValue()}:{name} (verify-exception: {ex.Message})")
                            Continue For
                        End Try

                        If p Is Nothing Then
                            skipCnt += 1
                            skipItems.Add($"{fi.Id.IntValue()}:{name} (child-inst-missing)")
                            Continue For
                        End If

                        Dim associated As FamilyParameter = Nothing
                        Try
                            associated = fm.GetAssociatedFamilyParameter(p)
                        Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                            failCnt += 1
                            fails.Add($"{fi.Id.IntValue()}:{name} (verify-exception: {ex.Message})")
                            Continue For
                        Catch ex As Exception
                            failCnt += 1
                            fails.Add($"{fi.Id.IntValue()}:{name} (verify-exception: {ex.Message})")
                            Continue For
                        End Try

                        If associated IsNot Nothing AndAlso associated.Id = hostParam.Id Then
                            okCnt += 1
                        Else
                            failCnt += 1
                            fails.Add($"{fi.Id.IntValue()}:{name} (not-associated)")
                        End If

                    Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                        failCnt += 1
                        fails.Add($"{fi.Id.IntValue()}:{name} (verify-exception: {ex.Message})")
                        Continue For
                    Catch ex As Exception
                        failCnt += 1
                        fails.Add($"{fi.Id.IntValue()}:{name} (verify-exception: {ex.Message})")
                        Continue For
                    End Try
                Next
            Next

            Dim r As New VerifyResult
            r.Ok = okCnt : r.Fail = failCnt : r.FailItems = fails
            r.Skip = skipCnt : r.SkipItems = skipItems
            Return r
        End Function

        Private Shared Function TryAssociateAcrossFamilyTypes(hostDoc As Document,
                                                              fm As FamilyManager,
                                                              fi As FamilyInstance,
                                                              paramName As String,
                                                              hostParam As FamilyParameter) As TypeDrivenAssociationResult
            Dim result As New TypeDrivenAssociationResult With {.Applicable = False, .Success = False}
            If hostDoc Is Nothing OrElse fm Is Nothing OrElse fi Is Nothing OrElse hostParam Is Nothing Then Return result

            If Not HasNestedTypeLabelAssociation(hostDoc, fm, fi) Then Return result
            result.Applicable = True

            Dim originalType As FamilyType = Nothing
            Try : originalType = fm.CurrentType : Catch : originalType = Nothing : End Try

            Try
                For Each famType As FamilyType In fm.Types
                    If famType Is Nothing Then Continue For

                    Try
                        fm.CurrentType = famType
                        hostDoc.Regenerate()
                    Catch
                        Continue For
                    End Try

                    Dim p As Parameter = Nothing
                    Try
                        p = TryGetElementParameterByName(fi, paramName)
                    Catch
                        p = Nothing
                    End Try
                    If p Is Nothing Then Continue For

                    Try
                        Dim associated = fm.GetAssociatedFamilyParameter(p)
                        If associated IsNot Nothing AndAlso associated.Id = hostParam.Id Then
                            result.Success = True
                            Return result
                        End If
                    Catch
                    End Try

                    Try
                        If (Not p.IsReadOnly) AndAlso fm.CanElementParameterBeAssociated(p) Then
                            fm.AssociateElementParameterToFamilyParameter(p, hostParam)
                            result.Success = True
                            Return result
                        End If
                    Catch
                    End Try
                Next
            Finally
                Try
                    If originalType IsNot Nothing Then fm.CurrentType = originalType
                Catch
                End Try
                Try : hostDoc.Regenerate() : Catch : End Try
            End Try

            Return result
        End Function

        Private Shared Function CollectChildFamiliesForGraph(hostDoc As Document,
                                                             fm As FamilyManager,
                                                             fi As FamilyInstance,
                                                             excludeDummy As Boolean,
                                                             ByRef dummyExcludedCount As Integer) As HashSet(Of String)
            Dim names As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If hostDoc Is Nothing OrElse fi Is Nothing Then Return names

            Dim currentChild As Family = TryGetNestedChildFamily(hostDoc, fi)
            AddGraphChildFamily(hostDoc, currentChild, excludeDummy, dummyExcludedCount, names)

            If fm Is Nothing Then Return names

            Dim labelKeys As HashSet(Of String) = GetNestedTypeLabelAssociationKeys(hostDoc, fm, fi)
            If labelKeys Is Nothing OrElse labelKeys.Count = 0 Then Return names

            Dim originalType As FamilyType = Nothing
            Try : originalType = fm.CurrentType : Catch : originalType = Nothing : End Try

            Try
                For Each famType As FamilyType In fm.Types
                    If famType Is Nothing Then Continue For

                    Try
                        fm.CurrentType = famType
                        hostDoc.Regenerate()
                    Catch
                        Continue For
                    End Try

                    CollectMatchingChildFamiliesForCurrentType(hostDoc,
                                                               fm,
                                                               labelKeys,
                                                               excludeDummy,
                                                               dummyExcludedCount,
                                                               names)
                Next
            Finally
                Try
                    If originalType IsNot Nothing Then fm.CurrentType = originalType
                Catch
                End Try
                Try : hostDoc.Regenerate() : Catch : End Try
            End Try

            Return names
        End Function

        Private Shared Function TryGetNestedChildFamily(hostDoc As Document,
                                                        fi As FamilyInstance) As Family
            If hostDoc Is Nothing OrElse fi Is Nothing Then Return Nothing

            Try
                Dim sym As FamilySymbol = TryCast(hostDoc.GetElement(fi.Symbol.Id), FamilySymbol)
                If sym Is Nothing Then Return Nothing
                Return sym.Family
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub CollectMatchingChildFamiliesForCurrentType(hostDoc As Document,
                                                                      fm As FamilyManager,
                                                                      labelKeys As HashSet(Of String),
                                                                      excludeDummy As Boolean,
                                                                      ByRef dummyExcludedCount As Integer,
                                                                      names As HashSet(Of String))
            If hostDoc Is Nothing OrElse fm Is Nothing OrElse labelKeys Is Nothing OrElse labelKeys.Count = 0 OrElse names Is Nothing Then Return

            Dim insts As IEnumerable(Of FamilyInstance) = Enumerable.Empty(Of FamilyInstance)()
            Try
                insts = New FilteredElementCollector(hostDoc).
                    OfClass(GetType(FamilyInstance)).
                    Cast(Of FamilyInstance)().
                    ToList()
            Catch
                Return
            End Try

            For Each candidate As FamilyInstance In insts
                If candidate Is Nothing Then Continue For

                Dim candidateKeys As HashSet(Of String) = GetNestedTypeLabelAssociationKeys(hostDoc, fm, candidate)
                If candidateKeys Is Nothing OrElse candidateKeys.Count = 0 Then Continue For
                If Not candidateKeys.Overlaps(labelKeys) Then Continue For

                Dim childFam As Family = TryGetNestedChildFamily(hostDoc, candidate)
                AddGraphChildFamily(hostDoc, childFam, excludeDummy, dummyExcludedCount, names)
            Next
        End Sub

        Private Shared Function GetNestedTypeLabelAssociationKeys(hostDoc As Document,
                                                                  fm As FamilyManager,
                                                                  fi As FamilyInstance) As HashSet(Of String)
            Dim keys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If hostDoc Is Nothing OrElse fm Is Nothing OrElse fi Is Nothing Then Return keys

            For Each p As Parameter In fi.Parameters
                If p Is Nothing Then Continue For

                Try
                    If p.StorageType <> StorageType.ElementId Then Continue For
                Catch
                    Continue For
                End Try

                Try
                    Dim associated = fm.GetAssociatedFamilyParameter(p)
                    If associated Is Nothing OrElse Not IsNestedTypeLabelFamilyParameter(associated) Then Continue For

                    Dim key As String = Nothing
                    Try
                        key = associated.Id.IntValue().ToString()
                    Catch
                        key = Nothing
                    End Try

                    If String.IsNullOrWhiteSpace(key) Then
                        Try
                            key = associated.Definition.Name
                        Catch
                            key = Nothing
                        End Try
                    End If

                    If Not String.IsNullOrWhiteSpace(key) Then
                        keys.Add(key)
                    End If
                Catch
                End Try
            Next

            Return keys
        End Function

        Private Shared Sub ReloadUpdatedChildrenIntoParentFamilyDoc(projDoc As Document,
                                                                    parentFamDoc As Document,
                                                                    parentFamilyName As String,
                                                                    childNames As HashSet(Of String),
                                                                    parentFails As List(Of String))
            If projDoc Is Nothing OrElse parentFamDoc Is Nothing OrElse childNames Is Nothing OrElse childNames.Count = 0 Then Return

            For Each childName In childNames
                If String.IsNullOrWhiteSpace(childName) Then Continue For

                Dim childProjFamily As Family = Nothing
                Dim childFamDoc As Document = Nothing
                Try
                    childProjFamily = FindProjectFamilyByName(projDoc, childName)
                    If childProjFamily Is Nothing OrElse Not childProjFamily.IsEditable Then
                        UI.Hub.UiBridgeExternalEvent.HostLog(
                            "debug",
                            $"[paramprop][child-reload-debug] stage=skip; host={parentFamilyName}; child={childName}; reason=project-family-missing")
                        Continue For
                    End If

                    childFamDoc = projDoc.EditFamily(childProjFamily)
                    TxnUtil.SafeLoadFamily(childFamDoc, parentFamDoc, $"KKY Reload child '{childName}' into '{parentFamilyName}'")
                    UI.Hub.UiBridgeExternalEvent.HostLog(
                        "debug",
                        $"[paramprop][child-reload-debug] stage=loaded; host={parentFamilyName}; child={childName}")
                Catch ex As Exception
                    If parentFails IsNot Nothing Then
                        parentFails.Add($"{parentFamilyName}: [ChildReload] {childName} - {ex.Message}")
                    End If
                    UI.Hub.UiBridgeExternalEvent.HostLog(
                        "debug",
                        $"[paramprop][child-reload-debug] stage=fail; host={parentFamilyName}; child={childName}; message={ex.Message}")
                Finally
                    If childFamDoc IsNot Nothing Then
                        Try : childFamDoc.Close(False) : Catch : End Try
                    End If
                End Try
            Next
        End Sub

        Private Shared Function FindProjectFamilyByName(projDoc As Document,
                                                        familyName As String) As Family
            If projDoc Is Nothing OrElse String.IsNullOrWhiteSpace(familyName) Then Return Nothing

            Try
                Return New FilteredElementCollector(projDoc).
                    OfClass(GetType(Family)).
                    Cast(Of Family)().
                    FirstOrDefault(Function(x) x IsNot Nothing AndAlso
                                              String.Equals(x.Name, familyName, StringComparison.OrdinalIgnoreCase))
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub SyncNestedChildDefinitionsInsideParentDoc(parentFamDoc As Document,
                                                                     parentFamilyName As String,
                                                                     childNames As HashSet(Of String),
                                                                     extDefs As IEnumerable(Of ExternalDefinition),
                                                                     chosenIsInstance As Boolean,
                                                                     chosenPG As BuiltInParameterGroup,
                                                                     parentFails As List(Of String))
            If parentFamDoc Is Nothing OrElse childNames Is Nothing OrElse childNames.Count = 0 OrElse extDefs Is Nothing Then Return

            For Each childName In childNames
                If String.IsNullOrWhiteSpace(childName) Then Continue For

                Dim nestedChildFamily As Family = Nothing
                Dim nestedChildDoc As Document = Nothing
                Try
                    nestedChildFamily = FindFamilyByNameInDocument(parentFamDoc, childName)
                    If nestedChildFamily Is Nothing OrElse Not nestedChildFamily.IsEditable Then
                        UI.Hub.UiBridgeExternalEvent.HostLog(
                            "debug",
                            $"[paramprop][nested-child-sync-debug] stage=skip; host={parentFamilyName}; child={childName}; reason=nested-family-missing")
                        Continue For
                    End If

                    nestedChildDoc = parentFamDoc.EditFamily(nestedChildFamily)

                    Dim addedAny As Boolean = False
                    For Each extDef In extDefs
                        If extDef Is Nothing Then Continue For
                        Dim res = EnsureSharedParamInFamily(nestedChildDoc, extDef, chosenIsInstance, chosenPG)
                        If res.Added Then addedAny = True
                        If Not String.IsNullOrWhiteSpace(res.ErrorMessage) Then
                            If parentFails IsNot Nothing Then
                                parentFails.Add($"{parentFamilyName}: [NestedChildSync] {childName} / {extDef.Name} - {res.ErrorMessage}")
                            End If
                        End If
                    Next

                    Try : nestedChildDoc.Regenerate() : Catch : End Try
                    TxnUtil.SafeLoadFamily(nestedChildDoc, parentFamDoc, $"KKY Sync nested child '{childName}' in '{parentFamilyName}'")

                    UI.Hub.UiBridgeExternalEvent.HostLog(
                        "debug",
                        $"[paramprop][nested-child-sync-debug] stage=loaded; host={parentFamilyName}; child={childName}; addedAny={addedAny}")
                Catch ex As Exception
                    If parentFails IsNot Nothing Then
                        parentFails.Add($"{parentFamilyName}: [NestedChildSync] {childName} - {ex.Message}")
                    End If
                    UI.Hub.UiBridgeExternalEvent.HostLog(
                        "debug",
                        $"[paramprop][nested-child-sync-debug] stage=fail; host={parentFamilyName}; child={childName}; message={ex.Message}")
                Finally
                    If nestedChildDoc IsNot Nothing Then
                        Try : nestedChildDoc.Close(False) : Catch : End Try
                    End If
                End Try
            Next
        End Sub

        Private Shared Function FindFamilyByNameInDocument(ownerDoc As Document,
                                                           familyName As String) As Family
            If ownerDoc Is Nothing OrElse String.IsNullOrWhiteSpace(familyName) Then Return Nothing

            Try
                Return New FilteredElementCollector(ownerDoc).
                    OfClass(GetType(Family)).
                    Cast(Of Family)().
                    FirstOrDefault(Function(x) x IsNot Nothing AndAlso
                                              String.Equals(x.Name, familyName, StringComparison.OrdinalIgnoreCase))
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub AddGraphChildFamily(hostDoc As Document,
                                               childFam As Family,
                                               excludeDummy As Boolean,
                                               ByRef dummyExcludedCount As Integer,
                                               names As HashSet(Of String))
            If names Is Nothing OrElse childFam Is Nothing Then Return
            If Not childFam.IsEditable Then Return
            If IsAnnotationFamily(childFam) Then Return

            If excludeDummy AndAlso childFam.Name.IndexOf("Dummy", StringComparison.OrdinalIgnoreCase) >= 0 Then
                dummyExcludedCount += 1
                Return
            End If

            Dim isShared As Boolean = False
            Try
                isShared = SharedFamilyHelper.IsFamilyShared(hostDoc, childFam)
            Catch
                isShared = False
            End Try
            If Not isShared Then Return

            names.Add(childFam.Name)
        End Sub

        Private Shared Function VerifyAssociationAcrossFamilyTypes(hostDoc As Document,
                                                                   fm As FamilyManager,
                                                                   fi As FamilyInstance,
                                                                   paramName As String,
                                                                   hostParam As FamilyParameter) As TypeDrivenAssociationResult
            Dim result As New TypeDrivenAssociationResult With {.Applicable = False, .Success = False}
            If hostDoc Is Nothing OrElse fm Is Nothing OrElse fi Is Nothing OrElse hostParam Is Nothing Then Return result

            If Not HasNestedTypeLabelAssociation(hostDoc, fm, fi) Then Return result
            result.Applicable = True

            Dim originalType As FamilyType = Nothing
            Try : originalType = fm.CurrentType : Catch : originalType = Nothing : End Try

            Try
                For Each famType As FamilyType In fm.Types
                    If famType Is Nothing Then Continue For

                    Try
                        fm.CurrentType = famType
                        hostDoc.Regenerate()
                    Catch
                        Continue For
                    End Try

                    Dim p As Parameter = Nothing
                    Try
                        p = TryGetElementParameterByName(fi, paramName)
                    Catch
                        p = Nothing
                    End Try
                    If p Is Nothing Then Continue For

                    Try
                        Dim associated = fm.GetAssociatedFamilyParameter(p)
                        If associated IsNot Nothing AndAlso associated.Id = hostParam.Id Then
                            result.Success = True
                            Return result
                        End If
                    Catch
                    End Try
                Next
            Finally
                Try
                    If originalType IsNot Nothing Then fm.CurrentType = originalType
                Catch
                End Try
                Try : hostDoc.Regenerate() : Catch : End Try
            End Try

            Return result
        End Function

        Private Shared Function HasNestedTypeLabelAssociation(hostDoc As Document,
                                                              fm As FamilyManager,
                                                              fi As FamilyInstance) As Boolean
            Return GetNestedTypeLabelAssociationKeys(hostDoc, fm, fi).Count > 0
        End Function

        Private Shared Function IsNestedTypeLabelFamilyParameter(param As FamilyParameter) As Boolean
            If param Is Nothing OrElse param.Definition Is Nothing Then Return False

#If REVIT2025 Then
            Try
                Dim dataType As ForgeTypeId = param.Definition.GetDataType()
                If dataType Is Nothing Then Return False
                Return Category.IsBuiltInCategory(dataType)
            Catch
                Return False
            End Try
#Else
            Return True
#End If
        End Function

        Private Shared Sub TraceAssociateSkipDebug(hostDoc As Document,
                                                   hostFamilyName As String,
                                                   childFamily As Family,
                                                   fi As FamilyInstance,
                                                   paramName As String,
                                                   hostParam As FamilyParameter,
                                                   reason As String,
                                                   childParam As Parameter,
                                                   fm As FamilyManager)
            Try
                Dim childFamilyName As String = If(childFamily Is Nothing, "<null>", childFamily.Name)
                Dim childParamFound As Boolean = (childParam IsNot Nothing)
                Dim childParamReadOnly As String = "<unknown>"
                Dim childParamAssociable As String = "<unknown>"
                Dim childParamStorage As String = "<unknown>"
                Dim labelDriven As Boolean = False

                If childParam IsNot Nothing Then
                    Try : childParamReadOnly = childParam.IsReadOnly.ToString() : Catch : End Try
                    Try : childParamAssociable = fm.CanElementParameterBeAssociated(childParam).ToString() : Catch : End Try
                    Try : childParamStorage = childParam.StorageType.ToString() : Catch : End Try
                End If

                Try : labelDriven = HasNestedTypeLabelAssociation(hostDoc, fm, fi) : Catch : labelDriven = False : End Try

                UI.Hub.UiBridgeExternalEvent.HostLog(
                    "debug",
                    $"[paramprop][associate-skip-debug] host={hostFamilyName}; child={childFamilyName}; instanceId={If(fi Is Nothing, -1, fi.Id.IntValue())}; param={paramName}; reason={reason}; childParamFound={childParamFound}; childParamReadOnly={childParamReadOnly}; childParamAssociable={childParamAssociable}; childParamStorage={childParamStorage}; labelDriven={labelDriven}; hostParamFound={hostParam IsNot Nothing}")
            Catch
            End Try
        End Sub

        Private Shared Function TryGetElementParameterByName(elem As Element,
                                                             paramName As String) As Parameter
            Dim ps As IList(Of Parameter) = elem.GetParameters(paramName)
            If ps IsNot Nothing AndAlso ps.Count > 0 Then Return ps(0)

            For Each p As Parameter In elem.Parameters
                Dim dn As Definition = p.Definition
                If dn IsNot Nothing AndAlso
                   String.Equals(dn.Name, paramName, StringComparison.OrdinalIgnoreCase) Then
                    Return p
                End If
            Next

            Return Nothing
        End Function

        Private Shared Function IsAnnotationFamily(fam As Family) As Boolean
            If fam Is Nothing Then Return False
            Dim cat = fam.FamilyCategory
            If cat Is Nothing Then Return False
            Return cat.CategoryType = CategoryType.Annotation
        End Function

        Private Shared Function ShouldSkipDummy(excludeDummy As Boolean,
                                                fam As Family) As Boolean
            If Not excludeDummy OrElse fam Is Nothing Then Return False
            Return fam.Name.IndexOf("Dummy", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Private Shared Function IsNoTxnNoise(msg As String) As Boolean
            If String.IsNullOrEmpty(msg) Then Return False
            Dim key As String = "Modification of the document is forbidden"
            Dim key2 As String = "no open transaction"
            Dim key3 As String = "Creation was undone"
            Return (msg.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0) OrElse
                   (msg.IndexOf(key2, StringComparison.OrdinalIgnoreCase) >= 0) OrElse
                   (msg.IndexOf(key3, StringComparison.OrdinalIgnoreCase) >= 0)
        End Function


        Private Shared Function NormalizeParamName(name As String) As String
            If name Is Nothing Then Return String.Empty
            Dim s As String = name.Trim()

            ' Remove wrapping quotes if any (e.g. "\"NAME\"")
            If s.Length >= 2 AndAlso s(0) = """"c AndAlso s(s.Length - 1) = """"c Then
                s = s.Substring(1, s.Length - 2).Trim()
            End If

            ' Normalize whitespace (tabs/NBSP -> space, collapse repeats)
            Dim sb As New StringBuilder(s.Length)
            Dim prevSpace As Boolean = False
            For Each ch As Char In s
                Dim isWs As Boolean = Char.IsWhiteSpace(ch) OrElse ch = ChrW(&HA0) ' NBSP
                If isWs Then
                    If Not prevSpace Then
                        sb.Append(" "c)
                        prevSpace = True
                    End If
                Else
                    sb.Append(ch)
                    prevSpace = False
                End If
            Next

            Return sb.ToString().Trim()
        End Function
        Private Shared Function NormalizeFolderPath(path As String) As String
            Dim s As String = If(path, String.Empty).Trim()
            If s.Length >= 2 AndAlso s.StartsWith("""", StringComparison.Ordinal) AndAlso s.EndsWith("""", StringComparison.Ordinal) Then
                s = s.Substring(1, s.Length - 2).Trim()
            End If
            Return s
        End Function
        Private Shared Function ResolveDefinitions(app As Application,
                                                   names As IEnumerable(Of String),
                                                   guids As IEnumerable(Of String)) As List(Of ExternalDefinition)
            Dim list As New List(Of ExternalDefinition)()
            If app Is Nothing Then Return list

            Dim reqNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If names IsNot Nothing Then
                For Each n In names
                    Dim nm As String = NormalizeParamName(If(n, String.Empty))
                    If Not String.IsNullOrWhiteSpace(nm) Then reqNames.Add(nm)
                Next
            End If

            Dim reqGuids As New HashSet(Of Guid)()
            If guids IsNot Nothing Then
                For Each g In guids
                    Dim gs As String = NormalizeParamName(If(g, String.Empty))
                    Dim id As Guid
                    If Guid.TryParse(gs, id) Then reqGuids.Add(id)
                Next
            End If

            If reqNames.Count = 0 AndAlso reqGuids.Count = 0 Then Return list

            Dim defFile = app.OpenSharedParameterFile()
            If defFile Is Nothing Then Return list

            Dim added As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each grp As DefinitionGroup In defFile.Groups
                If grp Is Nothing Then Continue For
                For Each defn As Definition In grp.Definitions
                    If defn Is Nothing Then Continue For
                    Dim ed = TryCast(defn, ExternalDefinition)
                    If ed Is Nothing Then Continue For

                    Dim hit As Boolean = False

                    If reqGuids.Count > 0 Then
                        Try
                            hit = reqGuids.Contains(ed.GUID)
                        Catch
                            hit = False
                        End Try
                    End If

                    If Not hit AndAlso reqNames.Count > 0 Then
                        Dim key As String = NormalizeParamName(ed.Name)
                        hit = reqNames.Contains(key)
                    End If

                    If hit Then
                        ' De-dupe by GUID when possible
                        Dim keyId As String = ""
                        Try : keyId = ed.GUID.ToString("D") : Catch : keyId = ed.Name : End Try
                        If Not added.Contains(keyId) Then
                            added.Add(keyId)
                            list.Add(ed)
                        End If
                    End If
                Next
            Next

            Return list
        End Function

        Private Shared Function BuildGroupOptions() As List(Of ParameterGroupOption)
#If REVIT2025 Then
            Dim items As New List(Of ParameterGroupOption)()
            For Each opt In BuiltInParameterGroupCompat.GetGroupOptions()
                If opt Is Nothing Then Continue For
                items.Add(New ParameterGroupOption With {
                    .Id = opt.Id,
                    .Key = opt.Key,
                    .Name = If(String.IsNullOrWhiteSpace(opt.Label), opt.Key, opt.Label)
                })
            Next

            Return items
#Else
            Dim items As New List(Of ParameterGroupOption)()
            Dim preferred As BuiltInParameterGroup() = {
                BuiltInParameterGroup.PG_TEXT,
                BuiltInParameterGroup.PG_IDENTITY_DATA,
                BuiltInParameterGroup.PG_DATA,
                BuiltInParameterGroup.PG_CONSTRAINTS
            }

            Dim added As New HashSet(Of Integer)()
            For Each pg In preferred
                items.Add(New ParameterGroupOption With {
                    .Id = CInt(pg),
                    .Key = pg.ToString(),
                    .Name = pg.ToString()
                })
                added.Add(CInt(pg))
            Next

            For Each pg As BuiltInParameterGroup In [Enum].GetValues(GetType(BuiltInParameterGroup))
                If Not added.Contains(CInt(pg)) Then
                    items.Add(New ParameterGroupOption With {
                        .Id = CInt(pg),
                        .Key = pg.ToString(),
                        .Name = pg.ToString()
                    })
                End If
            Next

            Return items
#End If
        End Function

        Private Shared Function BuildDetails(scanFails As IEnumerable(Of String),
                                             skips As IEnumerable(Of String),
                                             parentFails As IEnumerable(Of String),
                                             childFails As IEnumerable(Of String)) As List(Of SharedParamDetailRow)
            Dim rows As New List(Of SharedParamDetailRow)()
            AddRows(rows, scanFails, "ScanFail")
            AddRows(rows, skips, "Skip")
            AddRows(rows, parentFails, "Error")
            AddRows(rows, childFails, "ChildError")
            Return rows
        End Function

        Private Shared Sub AddRows(rows As List(Of SharedParamDetailRow),
                                   items As IEnumerable(Of String),
                                   kind As String)
            If rows Is Nothing OrElse items Is Nothing Then Return
            For Each s In items
                If String.IsNullOrWhiteSpace(s) Then Continue For
                Dim fam As String = s
                Dim detail As String = String.Empty
                Dim parts = s.Split(New Char() {":"c}, 2)
                If parts.Length = 2 Then
                    fam = parts(0).Trim()
                    detail = parts(1).Trim()
                End If
                rows.Add(New SharedParamDetailRow With {
                    .Kind = kind,
                    .Family = fam,
                    .Detail = detail
                })
            Next
        End Sub

    End Class

End Namespace
