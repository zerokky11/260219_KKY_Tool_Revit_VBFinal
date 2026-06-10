using System.Collections.Generic;
using System.Runtime.Serialization;

[DataContract]
public sealed class FamilyBrowserDeploymentBootstrapProfileIndex
{
	[DataMember(Name = "version", EmitDefaultValue = false)]
	public string Version { get; set; }

	[DataMember(Name = "defaultProfileId", EmitDefaultValue = false)]
	public string DefaultProfileId { get; set; }

	[DataMember(Name = "profiles", EmitDefaultValue = false)]
	public List<FamilyBrowserDeploymentBootstrapProfile> Profiles { get; set; }

	[DataMember(Name = "projectRules", EmitDefaultValue = false)]
	public List<FamilyBrowserDeploymentBootstrapProjectRule> ProjectRules { get; set; }

	public FamilyBrowserDeploymentBootstrapProfileIndex()
	{
		Version = string.Empty;
		DefaultProfileId = string.Empty;
		Profiles = new List<FamilyBrowserDeploymentBootstrapProfile>();
		ProjectRules = new List<FamilyBrowserDeploymentBootstrapProjectRule>();
	}
}
