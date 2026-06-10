using System.Runtime.Serialization;

[DataContract]
public sealed class FamilyBrowserStandardListJsonEntry
{
	[DataMember]
	public int RowNumber { get; set; }

	[DataMember]
	public string Discipline { get; set; }

	[DataMember]
	public string Category { get; set; }

	[DataMember]
	public string Family { get; set; }

	[DataMember]
	public string TypeName { get; set; }

	[DataMember]
	public string Notes { get; set; }

	public FamilyBrowserStandardListJsonEntry()
	{
		Discipline = string.Empty;
		Category = string.Empty;
		Family = string.Empty;
		TypeName = string.Empty;
		Notes = string.Empty;
	}
}
