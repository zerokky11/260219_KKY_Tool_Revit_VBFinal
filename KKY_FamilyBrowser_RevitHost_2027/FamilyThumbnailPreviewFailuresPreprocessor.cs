using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyThumbnailPreviewFailuresPreprocessor : IFailuresPreprocessor
{
	public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
	{
		//IL_0004: Unknown result type (might be due to invalid IL or missing references)
		//IL_006b: Unknown result type (might be due to invalid IL or missing references)
		//IL_006c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Invalid comparison between Unknown and I4
		if (failuresAccessor == null)
		{
			return (FailureProcessingResult)0;
		}
		try
		{
			IList<FailureMessageAccessor> messages = failuresAccessor.GetFailureMessages();
			if (messages == null)
			{
				return (FailureProcessingResult)0;
			}
			foreach (FailureMessageAccessor message in messages)
			{
				if (message != null && (int)message.GetSeverity() == 1)
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
		return (FailureProcessingResult)0;
	}
}
