using System.Collections.Generic;
using System.Runtime.Serialization;

[DataContract]
public sealed class FamilyBrowserBootstrapStandardLibrary
{
	[DataMember(Name = "discipline", EmitDefaultValue = false)]
	public string Discipline { get; set; }

	[DataMember(Name = "displayName", EmitDefaultValue = false)]
	public string DisplayName { get; set; }

	[DataMember(Name = "standardRvtPath", EmitDefaultValue = false)]
	public string StandardRvtPath { get; set; }

	[DataMember(Name = "standardRvtPathCandidates", EmitDefaultValue = false)]
	public List<string> StandardRvtPathCandidates { get; set; }

	[DataMember(Name = "registrationPath", EmitDefaultValue = false)]
	public string RegistrationPath { get; set; }

	[DataMember(Name = "registrationPathCandidates", EmitDefaultValue = false)]
	public List<string> RegistrationPathCandidates { get; set; }

	[DataMember(Name = "snapshotPath", EmitDefaultValue = false)]
	public string SnapshotPath { get; set; }

	[DataMember(Name = "snapshotPathCandidates", EmitDefaultValue = false)]
	public List<string> SnapshotPathCandidates { get; set; }

	[DataMember(Name = "standardListPath", EmitDefaultValue = false)]
	public string StandardListPath { get; set; }

	[DataMember(Name = "standardListPathCandidates", EmitDefaultValue = false)]
	public List<string> StandardListPathCandidates { get; set; }

	[DataMember(Name = "standardListSheetName", EmitDefaultValue = false)]
	public string StandardListSheetName { get; set; }

	[DataMember(Name = "sourceId", EmitDefaultValue = false)]
	public string SourceId { get; set; }

	[DataMember(Name = "disabled", EmitDefaultValue = false)]
	public bool Disabled { get; set; }

	public FamilyBrowserBootstrapStandardLibrary()
	{
		Discipline = string.Empty;
		DisplayName = string.Empty;
		StandardRvtPath = string.Empty;
		StandardRvtPathCandidates = new List<string>();
		RegistrationPath = string.Empty;
		RegistrationPathCandidates = new List<string>();
		SnapshotPath = string.Empty;
		SnapshotPathCandidates = new List<string>();
		StandardListPath = string.Empty;
		StandardListPathCandidates = new List<string>();
		StandardListSheetName = string.Empty;
		SourceId = string.Empty;
		Disabled = false;
	}
}
