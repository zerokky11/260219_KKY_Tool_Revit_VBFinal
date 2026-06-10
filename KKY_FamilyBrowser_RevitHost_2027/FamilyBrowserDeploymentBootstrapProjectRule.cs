using System.Runtime.Serialization;

[DataContract]
public sealed class FamilyBrowserDeploymentBootstrapProjectRule
{
	[DataMember(Name = "profileId", EmitDefaultValue = false)]
	public string ProfileId { get; set; }

	[DataMember(Name = "projectNameContains", EmitDefaultValue = false)]
	public string ProjectNameContains { get; set; }

	[DataMember(Name = "modelPathContains", EmitDefaultValue = false)]
	public string ModelPathContains { get; set; }

	[DataMember(Name = "centralPathContains", EmitDefaultValue = false)]
	public string CentralPathContains { get; set; }

	[DataMember(Name = "matchMode", EmitDefaultValue = false)]
	public string MatchMode { get; set; }

	[DataMember(Name = "matchValue", EmitDefaultValue = false)]
	public string MatchValue { get; set; }

	[DataMember(Name = "priority", EmitDefaultValue = false)]
	public int Priority { get; set; }

	[DataMember(Name = "disabled", EmitDefaultValue = false)]
	public bool Disabled { get; set; }

	public FamilyBrowserDeploymentBootstrapProjectRule()
	{
		ProfileId = string.Empty;
		ProjectNameContains = string.Empty;
		ModelPathContains = string.Empty;
		CentralPathContains = string.Empty;
		MatchMode = string.Empty;
		MatchValue = string.Empty;
		Priority = 0;
		Disabled = false;
	}
}
