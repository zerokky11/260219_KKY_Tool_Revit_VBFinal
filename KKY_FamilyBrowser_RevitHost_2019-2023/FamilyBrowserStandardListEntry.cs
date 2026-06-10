public sealed class FamilyBrowserStandardListEntry
{
	public int RowNumber { get; set; }

	public string Discipline { get; set; }

	public string Category { get; set; }

	public string Family { get; set; }

	public string TypeName { get; set; }

	public string Notes { get; set; }

	public FamilyBrowserStandardListEntry()
	{
		Discipline = string.Empty;
		Category = string.Empty;
		Family = string.Empty;
		TypeName = string.Empty;
		Notes = string.Empty;
	}
}
