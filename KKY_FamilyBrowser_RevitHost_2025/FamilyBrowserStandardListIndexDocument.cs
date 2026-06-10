using System.Collections.Generic;
using System.Runtime.Serialization;

[DataContract]
public sealed class FamilyBrowserStandardListIndexDocument
{
	[DataMember]
	public int SchemaVersion { get; set; }

	[DataMember]
	public string GeneratedAtUtc { get; set; }

	[DataMember]
	public string GeneratedBy { get; set; }

	[DataMember]
	public List<FamilyBrowserStandardListIndexItem> Items { get; set; }

	public FamilyBrowserStandardListIndexDocument()
	{
		SchemaVersion = 1;
		GeneratedAtUtc = string.Empty;
		GeneratedBy = string.Empty;
		Items = new List<FamilyBrowserStandardListIndexItem>();
	}
}
