public class StandardSystemTypeLayerSnapshotItem
{
	public int Index { get; set; }

	public string FunctionName { get; set; }

	public string MaterialName { get; set; }

	public string ThicknessDisplay { get; set; }

	public double ThicknessFeet { get; set; }

	public StandardSystemTypeLayerSnapshotItem()
	{
		FunctionName = string.Empty;
		MaterialName = string.Empty;
		ThicknessDisplay = string.Empty;
	}
}
