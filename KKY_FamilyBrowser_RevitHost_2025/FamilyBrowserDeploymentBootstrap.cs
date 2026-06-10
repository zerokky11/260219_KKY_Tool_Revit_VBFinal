using System.Collections.Generic;
using System.Runtime.Serialization;

[DataContract]
public sealed class FamilyBrowserDeploymentBootstrap
{
	[DataMember(Name = "version", EmitDefaultValue = false)]
	public string Version { get; set; }

	[DataMember(Name = "message", EmitDefaultValue = false)]
	public string Message { get; set; }

	[DataMember(Name = "disabled", EmitDefaultValue = false)]
	public bool Disabled { get; set; }

	[DataMember(Name = "refreshMinutes", EmitDefaultValue = false)]
	public int RefreshMinutes { get; set; }

	[DataMember(Name = "managedRootPath", EmitDefaultValue = false)]
	public string ManagedRootPath { get; set; }

	[DataMember(Name = "managedRootPathCandidates", EmitDefaultValue = false)]
	public List<string> ManagedRootPathCandidates { get; set; }

	[DataMember(Name = "managedPolicyPath", EmitDefaultValue = false)]
	public string ManagedPolicyPath { get; set; }

	[DataMember(Name = "managedPolicyPathCandidates", EmitDefaultValue = false)]
	public List<string> ManagedPolicyPathCandidates { get; set; }

	[DataMember(Name = "skipPolicyWrite", EmitDefaultValue = false)]
	public bool SkipPolicyWrite { get; set; }

	[DataMember(Name = "requestStore", EmitDefaultValue = false)]
	public FamilyBrowserBootstrapRequestStore RequestStore { get; set; }

	[DataMember(Name = "standardMode", EmitDefaultValue = false)]
	public string StandardMode { get; set; }

	[DataMember(Name = "standardLibraries", EmitDefaultValue = false)]
	public List<FamilyBrowserBootstrapStandardLibrary> StandardLibraries { get; set; }

	[DataMember(Name = "security", EmitDefaultValue = false)]
	public FamilyBrowserBootstrapSecurity Security { get; set; }

	public FamilyBrowserDeploymentBootstrap()
	{
		Version = string.Empty;
		Message = string.Empty;
		Disabled = false;
		RefreshMinutes = 30;
		ManagedRootPath = string.Empty;
		ManagedRootPathCandidates = new List<string>();
		ManagedPolicyPath = string.Empty;
		ManagedPolicyPathCandidates = new List<string>();
		SkipPolicyWrite = false;
		StandardMode = string.Empty;
		StandardLibraries = new List<FamilyBrowserBootstrapStandardLibrary>();
	}
}
