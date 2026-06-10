using System.Collections.Generic;

public class FamilyThumbnailGenerationResult
{
	public bool Success { get; set; }

	public string Message { get; set; }

	public List<string> Steps { get; set; }

	public bool ConnectorExtentsClamped { get; set; }

	public FamilyThumbnailGenerationResult()
	{
		Message = string.Empty;
		Steps = new List<string>();
	}
}
