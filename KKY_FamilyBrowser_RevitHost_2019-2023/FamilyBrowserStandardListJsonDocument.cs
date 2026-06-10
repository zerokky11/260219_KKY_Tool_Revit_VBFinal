using System.Collections.Generic;
using System.Runtime.Serialization;

[DataContract]
public sealed class FamilyBrowserStandardListJsonDocument
{
	[DataMember]
	public int SchemaVersion { get; set; }

	[DataMember]
	public string SourcePath { get; set; }

	[DataMember]
	public string SourceSheetName { get; set; }

	[DataMember]
	public string GeneratedAtUtc { get; set; }

	[DataMember]
	public string GeneratedBy { get; set; }

	[DataMember]
	public int RowCount { get; set; }

	[DataMember]
	public List<FamilyBrowserStandardListJsonEntry> Entries { get; set; }

	[DataMember(EmitDefaultValue = false)]
	public string BaselineCreatedAtUtc { get; set; }

	[DataMember(EmitDefaultValue = false)]
	public string BaselineCreatedBy { get; set; }

	[DataMember(EmitDefaultValue = false)]
	public string BaselineSourceSnapshotPath { get; set; }

	[DataMember(EmitDefaultValue = false)]
	public int BaselineSystemExclusionVersion { get; set; }

	[DataMember(EmitDefaultValue = false)]
	public List<FamilyBrowserStandardListJsonEntry> BaselineExcludedLoadableFamilies { get; set; }

	[DataMember(EmitDefaultValue = false)]
	public List<FamilyBrowserStandardListJsonEntry> BaselineExcludedSystemTypes { get; set; }

	public FamilyBrowserStandardListJsonDocument()
	{
		SchemaVersion = 1;
		SourcePath = string.Empty;
		SourceSheetName = string.Empty;
		GeneratedAtUtc = string.Empty;
		GeneratedBy = string.Empty;
		Entries = new List<FamilyBrowserStandardListJsonEntry>();
		BaselineCreatedAtUtc = string.Empty;
		BaselineCreatedBy = string.Empty;
		BaselineSourceSnapshotPath = string.Empty;
		BaselineExcludedLoadableFamilies = new List<FamilyBrowserStandardListJsonEntry>();
		BaselineExcludedSystemTypes = new List<FamilyBrowserStandardListJsonEntry>();
	}
}
