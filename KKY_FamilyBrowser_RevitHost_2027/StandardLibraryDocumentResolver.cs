using System;
using System.IO;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

public sealed class StandardLibraryDocumentResolver
{
	private StandardLibraryDocumentResolver()
	{
	}

	public static Document OpenRegisteredDocument(Application application, StandardLibraryRegistrationRecord registration, ref bool openedForCommand)
	{
		if (application == null)
		{
			throw new ArgumentNullException("application");
		}
		if (registration == null)
		{
			throw new ArgumentNullException("registration");
		}
		if (string.IsNullOrWhiteSpace(registration.ResolvedPath))
		{
			throw new InvalidOperationException(FamilyBrowserLanguageService.Text("The registered standard RVT path is empty.", "등록된 표준 RVT 경로가 비어 있습니다."));
		}
		string resolvedPath = Path.GetFullPath(registration.ResolvedPath);
		if (!File.Exists(resolvedPath))
		{
			throw new FileNotFoundException(FamilyBrowserLanguageService.Text("Registered standard RVT was not found.", "등록된 표준 RVT를 찾지 못했습니다."), resolvedPath);
		}
		Document standardDoc = FindOpenDocument(application, resolvedPath);
		if (standardDoc != null)
		{
			openedForCommand = false;
			return standardDoc;
		}
		openedForCommand = true;
		return application.OpenDocumentFile(resolvedPath);
	}

	private static Document FindOpenDocument(Application application, string resolvedPath)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_001a: Expected O, but got Unknown
		foreach (Document document in application.Documents)
		{
			Document doc = document;
			if (doc != null && !string.IsNullOrWhiteSpace(doc.PathName))
			{
				string candidatePath = string.Empty;
				try
				{
					candidatePath = Path.GetFullPath(doc.PathName);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
					continue;
				}
				if (string.Equals(candidatePath, resolvedPath, StringComparison.OrdinalIgnoreCase))
				{
					return doc;
				}
			}
		}
		return null;
	}
}
