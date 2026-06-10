using System.Collections.Generic;

public class StandardNestedLoadableFamilySnapshotItem
{
	public string FamilyName { get; set; }

	public string CategoryName { get; set; }

	public string CategoryId { get; set; }

	public string CategoryGroup { get; set; }

	public int TypeCount { get; set; }

	public List<string> TypeNames { get; set; }

	public bool IsShared { get; set; }

	public StandardNestedLoadableFamilySnapshotItem()
	{
		FamilyName = string.Empty;
		CategoryName = string.Empty;
		CategoryId = string.Empty;
		CategoryGroup = string.Empty;
		TypeNames = new List<string>();
	}
}
