using System;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserCommandError
{
	private FamilyBrowserCommandError()
	{
	}

	public static string ToExternalCommandMessage(string caption, Exception ex)
	{
		bool korean = IsKoreanUi();
		string ToExternalCommandMessage;
		try
		{
			string safeCaption = (string.IsNullOrWhiteSpace(caption) ? "Family Browser" : caption.Trim());
			string logPath = FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), safeCaption, ex);
			ToExternalCommandMessage = FamilyBrowserErrorHelp.Build(safeCaption, ex, logPath, korean).ToDialogMessage(korean);
		}
		catch (Exception ex2)
		{
			ProjectData.SetProjectError(ex2);
			Exception fallbackEx = ex2;
			string reason = ex?.Message ?? fallbackEx.Message;
			if (string.IsNullOrWhiteSpace(reason))
			{
				reason = (korean ? "알 수 없는 오류가 발생했습니다." : "An unknown error occurred.");
			}
			ToExternalCommandMessage = (korean ? ("Family Browser 작업을 완료하지 못했습니다." + Environment.NewLine + Environment.NewLine + "실패 이유" + Environment.NewLine + reason + Environment.NewLine + Environment.NewLine + "지금 할 일" + Environment.NewLine + "Revit 프로젝트와 표준 RVT/요청 저장소 연결 상태를 확인한 뒤 다시 실행하세요." + Environment.NewLine + Environment.NewLine + "관리자에게 전달할 정보" + Environment.NewLine + "오류 메시지와 실행한 작업명을 함께 전달하세요.") : ("Family Browser did not complete the action." + Environment.NewLine + Environment.NewLine + "Why It Failed" + Environment.NewLine + reason + Environment.NewLine + Environment.NewLine + "What To Do Now" + Environment.NewLine + "Check the Revit project, standard RVT, and request store connection, then try again." + Environment.NewLine + Environment.NewLine + "Send This To The Administrator" + Environment.NewLine + "Send the error message together with the action name."));
			ProjectData.ClearProjectError();
		}
		return ToExternalCommandMessage;
	}

	private static bool IsKoreanUi()
	{
		try
		{
			return !string.Equals(FamilyBrowserUserSettingsStore.LoadLanguageCode(), "en", StringComparison.OrdinalIgnoreCase);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return true;
	}
}
