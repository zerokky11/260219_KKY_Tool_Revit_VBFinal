using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyThumbnailPreviewFailuresPreprocessor : IFailuresPreprocessor
{
	public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
	{
		if (failuresAccessor == null)
		{
			return FailureProcessingResult.Continue;
		}
		try
		{
			IList<FailureMessageAccessor> messages = failuresAccessor.GetFailureMessages();
			if (messages == null)
			{
				return FailureProcessingResult.Continue;
			}
			foreach (FailureMessageAccessor message in messages)
			{
				if (message != null && message.GetSeverity() == FailureSeverity.Warning)
				{
					try
					{
						failuresAccessor.DeleteWarning(message);
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						ProjectData.ClearProjectError();
					}
				}
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		return FailureProcessingResult.Continue;
	}

	FailureProcessingResult IFailuresPreprocessor.PreprocessFailures(FailuresAccessor failuresAccessor)
	{
		//ILSpy generated this explicit interface implementation from .override directive in PreprocessFailures
		return this.PreprocessFailures(failuresAccessor);
	}
}
