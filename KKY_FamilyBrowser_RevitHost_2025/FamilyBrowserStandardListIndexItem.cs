using System.Runtime.Serialization;

[DataContract]
public sealed class FamilyBrowserStandardListIndexItem
{
	[DataMember]
	public string SlotKey { get; set; }

	[DataMember]
	public string Discipline { get; set; }

	[DataMember]
	public string DisplayName { get; set; }

	[DataMember]
	public string StandardListPath { get; set; }

	[DataMember]
	public string StandardRvtPath { get; set; }

	[DataMember]
	public string SourceId { get; set; }

	[DataMember]
	public bool Enabled { get; set; }

	public FamilyBrowserStandardListIndexItem()
	{
		SlotKey = string.Empty;
		Discipline = string.Empty;
		DisplayName = string.Empty;
		StandardListPath = string.Empty;
		StandardRvtPath = string.Empty;
		SourceId = string.Empty;
	}
}
