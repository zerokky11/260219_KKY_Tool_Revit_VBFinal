using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Microsoft.VisualBasic.CompilerServices;

public sealed class ProjectTrackingStoreService
{
	private static readonly Guid TrackingSchemaGuid = new Guid("21D8AA49-6D7E-468D-914A-0D5E73E52A11");

	private const string TrackingSchemaName = "KKYFamilyBrowserProjectTracking";

	private const string CatalogFieldName = "CatalogJson";

	private ProjectTrackingStoreService()
	{
	}

	public static ProjectTrackingCatalog Load(Document doc)
	{
		Schema trackingSchema = Schema.Lookup(TrackingSchemaGuid);
		if (trackingSchema == null)
		{
			return null;
		}
		DataStorage storage = FindStorage(doc, trackingSchema);
		if (storage == null)
		{
			return null;
		}
		Entity entity = ((Element)storage).GetEntity(trackingSchema);
		if (!entity.IsValid())
		{
			return null;
		}
		Field field = trackingSchema.GetField("CatalogJson");
		if (field == null)
		{
			return null;
		}
		string catalogJson = entity.Get<string>(field);
		if (string.IsNullOrWhiteSpace(catalogJson))
		{
			return null;
		}
		return DataContractJsonTextStore.Load<ProjectTrackingCatalog>(catalogJson);
	}

	public static void Save(Document doc, ProjectTrackingCatalog catalog)
	{
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Expected O, but got Unknown
		Schema trackingSchema = EnsureSchema();
		DataStorage storage = FindStorage(doc, trackingSchema);
		if (storage == null)
		{
			storage = DataStorage.Create(doc);
		}
		Entity entity = new Entity(trackingSchema);
		entity.Set<string>(trackingSchema.GetField("CatalogJson"), PlainJsonReportWriter.Serialize(catalog));
		((Element)storage).SetEntity(entity);
	}

	public static void MarkCurrentModelCheckRequired(Document doc, string currentUser, string reason, IEnumerable<ProjectTrackingDirtyItem> items)
	{
		if (doc != null)
		{
			try
			{
				ProjectTrackingDirtyMarker marker = new ProjectTrackingDirtyMarker
				{
					DetectedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
					User = (currentUser ?? string.Empty),
					DocumentTitle = (doc.Title ?? string.Empty),
					DocumentPath = ProjectSnapshotStore.ResolveProjectIdentityPath(doc),
					State = "NeedsAdminRestoreFromStandard",
					RequiredAction = "CurrentModelCheckRequired",
					Reason = (reason ?? string.Empty),
					Items = (items ?? new List<ProjectTrackingDirtyItem>()).ToList()
				};
				string path = BuildDirtyMarkerPath(doc);
				Directory.CreateDirectory(Path.GetDirectoryName(path));
				File.WriteAllText(path, PlainJsonReportWriter.Serialize(marker), Encoding.UTF8);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	public static ProjectTrackingDirtyMarker LoadCurrentModelCheckMarker(Document doc)
	{
		ProjectTrackingDirtyMarker LoadCurrentModelCheckMarker;
		if (doc == null)
		{
			LoadCurrentModelCheckMarker = null;
		}
		else
		{
			try
			{
				string markerPath = BuildDirtyMarkerPath(doc);
				LoadCurrentModelCheckMarker = (File.Exists(markerPath) ? DataContractJsonFileStore.Load<ProjectTrackingDirtyMarker>(markerPath) : null);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				LoadCurrentModelCheckMarker = null;
				ProjectData.ClearProjectError();
			}
		}
		return LoadCurrentModelCheckMarker;
	}

	public static void ClearCurrentModelCheckRequired(Document doc)
	{
		if (doc == null)
		{
			return;
		}
		try
		{
			string markerPath = BuildDirtyMarkerPath(doc);
			if (File.Exists(markerPath))
			{
				File.Delete(markerPath);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static Schema EnsureSchema()
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0030: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		Schema trackingSchema = Schema.Lookup(TrackingSchemaGuid);
		if (trackingSchema != null)
		{
			return trackingSchema;
		}
		SchemaBuilder val = new SchemaBuilder(TrackingSchemaGuid);
		val.SetSchemaName("KKYFamilyBrowserProjectTracking");
		val.SetReadAccessLevel((AccessLevel)1);
		val.SetWriteAccessLevel((AccessLevel)1);
		val.SetDocumentation("KKY Family Browser tracked project state catalog.");
		val.AddSimpleField("CatalogJson", typeof(string));
		return val.Finish();
	}

	private static DataStorage FindStorage(Document doc, Schema schema)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		foreach (Element item in new FilteredElementCollector(doc).OfClass(typeof(DataStorage)))
		{
			DataStorage storage = (DataStorage)(object)((item is DataStorage) ? item : null);
			if (storage != null && ((Element)storage).GetEntity(schema).IsValid())
			{
				return storage;
			}
		}
		return null;
	}

	private static string BuildDirtyMarkerPath(Document doc)
	{
		string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
		if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
		{
			return string.Empty;
		}
		return Path.Combine(Path.Combine(ProjectSnapshotStore.GetProjectHistoryFolder(workspaceRoot, doc), "ProjectTrackingDirty"), BuildDocumentHash(doc) + "-current-model-check-required.json");
	}

	private static string BuildDocumentHash(Document doc)
	{
		string key = ProjectSnapshotStore.ResolveProjectIdentityPath(doc).Trim();
		if (string.IsNullOrWhiteSpace(key))
		{
			key = (doc.Title ?? string.Empty) + "|" + doc.GetHashCode().ToString(CultureInfo.InvariantCulture);
		}
		using SHA256 sha = SHA256.Create();
		byte[] array = sha.ComputeHash(Encoding.UTF8.GetBytes(key));
		StringBuilder sb = new StringBuilder();
		byte[] array2 = array;
		foreach (byte value in array2)
		{
			sb.Append(value.ToString("x2", CultureInfo.InvariantCulture));
		}
		return sb.ToString();
	}
}
