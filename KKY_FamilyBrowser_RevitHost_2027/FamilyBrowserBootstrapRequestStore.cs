using System.Collections.Generic;
using System.Runtime.Serialization;

[DataContract]
public sealed class FamilyBrowserBootstrapRequestStore
{
	[DataMember(Name = "mode", EmitDefaultValue = false)]
	public string Mode { get; set; }

	[DataMember(Name = "path", EmitDefaultValue = false)]
	public string Path { get; set; }

	[DataMember(Name = "pathCandidates", EmitDefaultValue = false)]
	public List<string> PathCandidates { get; set; }

	[DataMember(Name = "endpoint", EmitDefaultValue = false)]
	public string Endpoint { get; set; }

	public FamilyBrowserBootstrapRequestStore()
	{
		Mode = string.Empty;
		Path = string.Empty;
		PathCandidates = new List<string>();
		Endpoint = string.Empty;
	}
}
