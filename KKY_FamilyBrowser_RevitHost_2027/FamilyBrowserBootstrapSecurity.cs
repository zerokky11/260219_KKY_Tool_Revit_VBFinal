using System.Collections.Generic;
using System.Runtime.Serialization;

[DataContract]
public sealed class FamilyBrowserBootstrapSecurity
{
	[DataMember(Name = "adminUsers", EmitDefaultValue = false)]
	public List<string> AdminUsers { get; set; }

	[DataMember(Name = "adminProfileKeywords", EmitDefaultValue = false)]
	public List<string> AdminProfileKeywords { get; set; }

	[DataMember(Name = "requestApproverUsers", EmitDefaultValue = false)]
	public List<string> RequestApproverUsers { get; set; }

	[DataMember(Name = "readOnlyUsers", EmitDefaultValue = false)]
	public List<string> ReadOnlyUsers { get; set; }

	public FamilyBrowserBootstrapSecurity()
	{
		AdminUsers = new List<string>();
		AdminProfileKeywords = new List<string>();
		RequestApproverUsers = new List<string>();
		ReadOnlyUsers = new List<string>();
	}
}
