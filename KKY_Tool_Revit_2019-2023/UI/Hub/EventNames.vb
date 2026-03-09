Option Explicit On
Option Strict On

Namespace UI.Hub

    Public NotInheritable Class EventNames
        Private Sub New()
        End Sub

        Public NotInheritable Class Ui
            Private Sub New()
            End Sub

            Public Const Ping As String = "ui:ping"
            Public Const QueryTopmost As String = "ui:query-topmost"
            Public Const SetTopmost As String = "ui:set-topmost"
            Public Const ToggleTopmost As String = "ui:toggle-topmost"
        End Class

        Public NotInheritable Class Host
            Private Sub New()
            End Sub

            Public Const Connected As String = "host:connected"
            Public Const Pong As String = "host:pong"
            Public Const Topmost As String = "host:topmost"
            Public Const DocChanged As String = "host:doc-changed"
            Public Const DocList As String = "host:doc-list"
            Public Const ErrorEvent As String = "host:error"
            Public Const Warn As String = "host:warn"
            Public Const Info As String = "host:info"
            Public Const Log As String = "host:log"
        End Class

        Public NotInheritable Class Feature
            Private Sub New()
            End Sub

            Public NotInheritable Class Ui
                Private Sub New()
                End Sub

                Public Const DupRun As String = "dup:run"
                Public Const DuplicateExport As String = "duplicate:export"
                Public Const DuplicateDelete As String = "duplicate:delete"
                Public Const DuplicateRestore As String = "duplicate:restore"
                Public Const DuplicateSelect As String = "duplicate:select"

                Public Const ConnectorRun As String = "connector:run"
                Public Const ConnectorSaveExcel As String = "connector:save-excel"

                Public Const ExportBrowseFolder As String = "export:browse-folder"
                Public Const ExportAddRvtFiles As String = "export:add-rvt-files"
                Public Const ExportPreview As String = "export:preview"
                Public Const ExportSaveExcel As String = "export:save-excel"

                Public Const ParamPropRun As String = "paramprop:run"
                Public Const SharedParamRun As String = "sharedparam:run"
                Public Const SharedParamList As String = "sharedparam:list"
                Public Const SharedParamStatus As String = "sharedparam:status"
                Public Const SharedParamExportExcel As String = "sharedparam:export-excel"
                Public Const ParamPropList As String = "paramprop:list"
                Public Const ParamPropStatus As String = "paramprop:status"
                Public Const ParamPropExportExcel As String = "paramprop:export-excel"
                Public Const SharedParamExport As String = "sharedparam:export"

                Public Const SharedParamBatchInit As String = "sharedparambatch:init"
                Public Const SharedParamBatchBrowseRvts As String = "sharedparambatch:browse-rvts"
                Public Const SharedParamBatchBrowseFolder As String = "sharedparambatch:browse-folder"
                Public Const SharedParamBatchRun As String = "sharedparambatch:run"
                Public Const SharedParamBatchExportExcel As String = "sharedparambatch:export-excel"
                Public Const SharedParamBatchOpenFolder As String = "sharedparambatch:open-folder"

                Public Const ExcelOpen As String = "excel:open"

                Public Const SegmentPmsRvtPickFiles As String = "segmentpms:rvt-pick-files"
                Public Const SegmentPmsRvtPickFolder As String = "segmentpms:rvt-pick-folder"
                Public Const SegmentPmsExtract As String = "segmentpms:extract"
                Public Const SegmentPmsLoadExtract As String = "segmentpms:load-extract"
                Public Const SegmentPmsSaveExtract As String = "segmentpms:save-extract"
                Public Const SegmentPmsRegisterPms As String = "segmentpms:register-pms"
                Public Const SegmentPmsPmsTemplate As String = "segmentpms:pms-template"
                Public Const SegmentPmsPrepareMapping As String = "segmentpms:prepare-mapping"
                Public Const SegmentPmsRun As String = "segmentpms:run"
                Public Const SegmentPmsSaveResult As String = "segmentpms:save-result"

                Public Const GuidAddFiles As String = "guid:add-files"
                Public Const GuidRun As String = "guid:run"
                Public Const GuidExport As String = "guid:export"
                Public Const GuidRequestFamilyDetail As String = "guid:request-family-detail"

                Public Const FamilyLinkInit As String = "familylink:init"
                Public Const FamilyLinkPickRvts As String = "familylink:pick-rvts"
                Public Const FamilyLinkRun As String = "familylink:run"
                Public Const FamilyLinkExport As String = "familylink:export"

                Public Const HubPickRvt As String = "hub:pick-rvt"
                Public Const HubMultiRun As String = "hub:multi-run"
                Public Const HubMultiExport As String = "hub:multi-export"
                Public Const HubMultiClear As String = "hub:multi-clear"

                Public Const CommonOptionsGet As String = "commonoptions:get"
                Public Const CommonOptionsSave As String = "commonoptions:save"
            End Class

            Public NotInheritable Class Host
                Private Sub New()
                End Sub

                Public Const DupList As String = "dup:list"
                Public Const DupExported As String = "dup:exported"
                Public Const ConnectorDone As String = "connector:done"
                Public Const ExportPreviewed As String = "export:previewed"
                Public Const ExportSaved As String = "export:saved"
                Public Const ParamPropDone As String = "paramprop:done"
                Public Const ParamPropReport As String = "paramprop:report"
                Public Const FamilyLinkSharedParams As String = "familylink:sharedparams"
                Public Const FamilyLinkRvtsPicked As String = "familylink:rvts-picked"
                Public Const FamilyLinkProgress As String = "familylink:progress"
                Public Const FamilyLinkResult As String = "familylink:result"
                Public Const FamilyLinkExported As String = "familylink:exported"
            End Class
        End Class
    End Class
End Namespace