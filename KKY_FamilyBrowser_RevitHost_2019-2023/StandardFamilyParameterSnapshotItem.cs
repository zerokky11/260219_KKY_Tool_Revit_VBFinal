public class StandardFamilyParameterSnapshotItem
{
	public string Scope { get; set; }

	public string TypeName { get; set; }

	public string Name { get; set; }

	public string StorageType { get; set; }

	public string ValuePreview { get; set; }

	public string Formula { get; set; }

	public bool IsInstance { get; set; }

	public bool IsReadOnly { get; set; }

	public bool IsShared { get; set; }

	public string ParameterId { get; set; }

	public string ExternalGuid { get; set; }

	public StandardFamilyParameterSnapshotItem()
	{
		Scope = string.Empty;
		TypeName = string.Empty;
		Name = string.Empty;
		StorageType = string.Empty;
		ValuePreview = string.Empty;
		Formula = string.Empty;
		ParameterId = string.Empty;
		ExternalGuid = string.Empty;
	}
}
