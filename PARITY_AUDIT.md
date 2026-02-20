# VB->C# Parity Audit

- Total pairs: 36
- HIGH risk: 8
- MEDIUM risk: 4

## Summary Table
|Risk|VB File|CS File|VB Lines|CS Lines|Line Ratio|VB Methods|CS Methods|Missing Methods|
|---|---|---:|---:|---:|---:|---:|---:|---:|
|HIGH|`Services/SegmentPmsCheckService.vb`|`Services/SegmentPmsCheckService.cs`|2623|241|0.09|82|13|72|
|HIGH|`UI/Hub/UiBridgeExternalEvent.Multi.vb`|`UI/Hub/UiBridgeExternalEvent.Multi.cs`|1075|156|0.15|51|5|46|
|HIGH|`Services/SharedParamBatchService.vb`|`Services/SharedParamBatchService.cs`|1614|476|0.29|51|15|44|
|HIGH|`Services/ParamPropagateService.vb`|`Services/ParamPropagateService.cs`|1975|137|0.07|45|4|41|
|HIGH|`UI/Hub/UiBridgeExternalEvent.Connector.vb`|`UI/Hub/UiBridgeExternalEvent.Connector.cs`|1189|254|0.21|52|18|40|
|HIGH|`Services/GuidAuditService.vb`|`Services/GuidAuditService.cs`|1261|186|0.15|35|6|34|
|HIGH|`UI/Hub/UiBridgeExternalEvent.SegmentPms.vb`|`UI/Hub/UiBridgeExternalEvent.SegmentPms.cs`|1014|190|0.19|36|18|22|
|HIGH|`Services/ConnectorDiagnosticsService.vb`|`Services/ConnectorDiagnosticsService.cs`|1301|569|0.44|38|22|20|
|MEDIUM|`UI/Hub/UiBridgeExternalEvent.Guid.vb`|`UI/Hub/UiBridgeExternalEvent.Guid.cs`|369|148|0.40|12|5|8|
|LOW|`UI/Hub/UiBridgeExternalEvent.Duplicate.vb`|`UI/Hub/UiBridgeExternalEvent.Duplicate.cs`|978|754|0.77|25|20|5|
|LOW|`UI/Hub/UiBridgeExternalEvent.Core.vb`|`UI/Hub/UiBridgeExternalEvent.Core.cs`|430|570|1.33|16|21|4|
|MEDIUM|`My Project/Settings.Designer.vb`|`My Project/Settings.Designer.cs`|73|33|0.45|1|0|1|
|LOW|`Exports/ConnectorExport.vb`|`Exports/ConnectorExport.cs`|47|55|1.17|3|3|0|
|LOW|`Exports/DuplicateExport.vb`|`Exports/DuplicateExport.cs`|167|196|1.17|10|10|0|
|LOW|`Exports/FamilyLinkAuditExport.vb`|`Exports/FamilyLinkAuditExport.cs`|151|169|1.12|7|7|0|
|LOW|`Exports/PointsExport.vb`|`Exports/PointsExport.cs`|29|29|1.00|2|2|0|
|MEDIUM|`Infrastructure/ElementIdCompat.vb`|`Infrastructure/ElementIdCompat.cs`|55|23|0.42|3|3|0|
|LOW|`Infrastructure/ExcelCore.vb`|`Infrastructure/ExcelCore.cs`|831|956|1.15|30|30|0|
|LOW|`Infrastructure/ExcelExportStyleRegistry.vb`|`Infrastructure/ExcelExportStyleRegistry.cs`|278|298|1.07|20|20|0|
|LOW|`Infrastructure/ExcelStyleHelper.vb`|`Infrastructure/ExcelStyleHelper.cs`|133|125|0.94|9|9|0|
|LOW|`Infrastructure/ResourceExtractor.vb`|`Infrastructure/ResourceExtractor.cs`|106|118|1.11|2|2|0|
|LOW|`Infrastructure/ResultTableFilter.vb`|`Infrastructure/ResultTableFilter.cs`|49|47|0.96|2|2|0|
|LOW|`My Project/Application.Designer.vb`|`My Project/Application.Designer.cs`|13|9|0.69|0|0|0|
|MEDIUM|`My Project/AssemblyInfo.vb`|`My Project/AssemblyInfo.cs`|32|15|0.47|0|0|0|
|LOW|`My Project/Resources.Designer.vb`|`My Project/Resources.Designer.cs`|62|40|0.65|0|0|0|
|LOW|`Services/DuplicateAnalysisService.vb`|`Services/DuplicateAnalysisService.cs`|234|230|0.98|7|7|0|
|LOW|`Services/ExportPointsService.vb`|`Services/ExportPointsService.cs`|335|366|1.09|17|17|0|
|LOW|`Services/FamilyLinkAuditService.vb`|`Services/FamilyLinkAuditService.cs`|565|634|1.12|15|15|0|
|LOW|`Services/HubCommonOptionsStorageService.vb`|`Services/HubCommonOptionsStorageService.cs`|67|77|1.15|3|3|0|
|LOW|`Services/SharedParameterStatusService.vb`|`Services/SharedParameterStatusService.cs`|166|187|1.13|3|3|0|
|LOW|`UI/Hub/ExcelProgressReporter.vb`|`UI/Hub/ExcelProgressReporter.cs`|84|97|1.15|3|3|0|
|LOW|`UI/Hub/HubHostWindow.vb`|`UI/Hub/HubHostWindow.cs`|296|390|1.32|15|15|0|
|LOW|`UI/Hub/UiBridgeExternalEvent.Export.vb`|`UI/Hub/UiBridgeExternalEvent.Export.cs`|583|683|1.17|31|31|0|
|LOW|`UI/Hub/UiBridgeExternalEvent.FamilyLinkAudit.vb`|`UI/Hub/UiBridgeExternalEvent.FamilyLinkAudit.cs`|337|408|1.21|12|13|0|
|LOW|`UI/Hub/UiBridgeExternalEvent.ParamProp.vb`|`UI/Hub/UiBridgeExternalEvent.ParamProp.cs`|210|237|1.13|6|6|0|
|LOW|`UI/Hub/UiBridgeExternalEvent.SharedParamBatch.vb`|`UI/Hub/UiBridgeExternalEvent.SharedParamBatch.cs`|214|286|1.34|7|8|0|

## Missing Method Samples (Top 12 per file)
- `Services/SegmentPmsCheckService.vb`: addcomparerow, adderror, addmissingmappingrows, applyworksetconfiguration, buildcomparetable, builderrortable, buildfiletable, buildmaptable, buildmetatable, buildopenoptions, buildpipetypeclassmap, buildpmstableskeleton
- `UI/Hub/UiBridgeExternalEvent.Multi.vb`: addemptymessagerow, anyfeatureenabled, appendmulticonnectorerror, appendmultirunitem, appendsegmentpmsrows, buildconnectorheaders, buildconnectortablefromrows, buildmappingsfromsuggestions, buildopenoptions, buildpointheaders, buildpointtable, buildtablefromrows
- `Services/SharedParamBatchService.vb`: addlog, applyallsharedparameterbindings, applyonesharedparameterbinding, buildavailablecategorymaps, buildcategorytree, buildstatusmap, clone, clonedeep, collectcategoryrecursive, collectrvtfiles, describecategoryref, findexternaldefinitionbyguid
- `Services/ParamPropagateService.vb`: addrows, builddetails, buildgroupoptions, buildrows, buildui, ensureparampropexportschema, ensuresharedparaminfamily, escapecsv, executecore, findfamilyidbyname, getparamtypestring, isannotationfamily
- `UI/Hub/UiBridgeExternalEvent.Connector.vb`: appendextrasforid1, applyfastcolumnwidths, buildbaseheaders, buildconnectorexportdatatable, buildemptyconnectorrows, buildextraheaders, buildheaders, buildreviewrows, clonerow, convertdistanceforui, countmismatches, createborderedstyle
- `Services/GuidAuditService.vb`: addallowedcategoryname, adddetailrow, addopenfailrow, appendnote, buildallowedcategorynameset, buildexceptionnotes, buildopenfailnotes, buildtargets, createdetachoptions, export, exportmulti, findopendocument
- `UI/Hub/UiBridgeExternalEvent.SegmentPms.vb`: addsheet, buildemptyrows, buildextractsummary, buildgrouppayload, buildpmsoptions, buildtablefromrows, coercerowstodictlist, datatabletoobjects, dictlisttodatatable, equals, getdictvalue, gethashcode
- `Services/ConnectorDiagnosticsService.vb`: atend, bucketkey, buildgrid, expect, expectvalue, findproximitycandidate, getextravalues, getparaminfo, makeoriginkey, makeoriginpairkey, nextis, parse
- `UI/Hub/UiBridgeExternalEvent.Guid.vb`: clonewithoutcolumn, ensurenorvtpath, filterfamilydetail, handleguidaddfolder, reportguidprogress, safeboolobj, safestrguid, shapetable
- `UI/Hub/UiBridgeExternalEvent.Duplicate.vb`: addhandler, btnbar, buildexcelsavedcontent, issystemdark, showexcelsaveddialog
- `UI/Hub/UiBridgeExternalEvent.Core.vb`: filterissuerowscopy, handleswitchdocument, hostlog, logautofitdecision
- `My Project/Settings.Designer.vb`: autosavesettings
