using System.Runtime.Serialization;

[DataContract]
public sealed class FamilyBrowserDeploymentBootstrapProfile
{
	[DataMember(Name = "id", EmitDefaultValue = false)]
	public string Id { get; set; }

	[DataMember(Name = "name", EmitDefaultValue = false)]
	public string Name { get; set; }

	[DataMember(Name = "description", EmitDefaultValue = false)]
	public string Description { get; set; }

	[DataMember(Name = "url", EmitDefaultValue = false)]
	public string Url { get; set; }

	[DataMember(Name = "disabled", EmitDefaultValue = false)]
	public bool Disabled { get; set; }

	public FamilyBrowserDeploymentBootstrapProfile()
	{
		Id = string.Empty;
		Name = string.Empty;
		Description = string.Empty;
		Url = string.Empty;
		Disabled = false;
	}
}
