using System.Text;

public sealed class FamilyBrowserFriendlyError
{
	public string Title { get; set; }

	public string Summary { get; set; }

	public string Cause { get; set; }

	public string UserAction { get; set; }

	public string AdminAction { get; set; }

	public string TechnicalDetail { get; set; }

	public string LogPath { get; set; }

	public string SupportCode { get; set; }

	public FamilyBrowserFriendlyError()
	{
		Title = string.Empty;
		Summary = string.Empty;
		Cause = string.Empty;
		UserAction = string.Empty;
		AdminAction = string.Empty;
		TechnicalDetail = string.Empty;
		LogPath = string.Empty;
		SupportCode = string.Empty;
	}

	public string ToDialogMessage(bool korean)
	{
		StringBuilder builder = new StringBuilder();
		builder.AppendLine(Title);
		builder.AppendLine();
		builder.AppendLine(korean ? "실패 이유" : "Why It Failed");
		builder.AppendLine(string.IsNullOrWhiteSpace(Cause) ? Summary : Cause);
		builder.AppendLine();
		builder.AppendLine(korean ? "지금 할 일" : "What To Do Now");
		builder.AppendLine(UserAction);
		builder.AppendLine();
		builder.AppendLine(korean ? "관리자에게 전달할 정보" : "Send This To The Administrator");
		builder.AppendLine(AdminAction);
		if (!string.IsNullOrWhiteSpace(LogPath))
		{
			builder.AppendLine((korean ? "로그: " : "Log: ") + LogPath);
		}
		builder.AppendLine((korean ? "지원 코드: " : "Support code: ") + SupportCode);
		if (!string.IsNullOrWhiteSpace(TechnicalDetail))
		{
			builder.AppendLine();
			builder.AppendLine(korean ? "기술 정보" : "Technical Detail");
			builder.AppendLine(TechnicalDetail);
		}
		return builder.ToString();
	}
}
