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
		Entity entity = storage.GetEntity(trackingSchema);
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
		Schema trackingSchema = EnsureSchema();
		DataStorage storage = FindStorage(doc, trackingSchema);
		if (storage == null)
		{
			storage = DataStorage.Create(doc);
		}
		Entity entity = new Entity(trackingSchema);
		entity.Set(trackingSchema.GetField("CatalogJson"), PlainJsonReportWriter.Serialize(catalog));
		storage.SetEntity(entity);
	}

	public static void MarkCurrentModelCheckRequired(Document doc, string currentUser, string reason, IEnumerable<ProjectTrackingDirtyItem> items)
	{
		if (doc != null)
		{
			try
			{
				string path = BuildDirtyMarkerPath(doc);
				if (string.IsNullOrWhiteSpace(path))
				{
					return;
				}
				Directory.CreateDirectory(Path.GetDirectoryName(path));
				using (FileStream mutationLock = AcquireDirtyMarkerMutationLock(path))
				{
					ProjectTrackingDirtyMarker existing = LoadMarkerCore(path);
				ProjectTrackingDirtyMarker marker = new ProjectTrackingDirtyMarker
				{
					MarkerId = Guid.NewGuid().ToString("N"),
					DetectedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
					User = (currentUser ?? string.Empty),
					DocumentTitle = (doc.Title ?? string.Empty),
					DocumentPath = ProjectSnapshotStore.ResolveProjectIdentityPath(doc),
					State = "NeedsAdminRestoreFromStandard",
					RequiredAction = "CurrentModelCheckRequired",
					Reason = MergeDirtyMarkerReasons(existing == null ? string.Empty : existing.Reason, reason),
					Items = MergeDirtyMarkerItems(existing == null ? null : existing.Items, items)
				};
					WriteMarkerAtomic(path, marker);
				}
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
				LoadCurrentModelCheckMarker = LoadMarkerCore(markerPath);
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

	public static bool ClearCurrentModelCheckRequired(Document doc, ProjectTrackingDirtyMarker expectedMarker)
	{
		if (doc == null || expectedMarker == null)
		{
			return false;
		}
		try
		{
			string markerPath = BuildDirtyMarkerPath(doc);
			if (string.IsNullOrWhiteSpace(markerPath))
			{
				return false;
			}
			using (FileStream mutationLock = AcquireDirtyMarkerMutationLock(markerPath))
			{
				ProjectTrackingDirtyMarker current = LoadMarkerCore(markerPath);
				if (!DirtyMarkerIdentityMatches(current, expectedMarker))
				{
					return false;
				}
				if (File.Exists(markerPath))
				{
					File.Delete(markerPath);
				}
				return true;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
			return false;
		}
	}

	private static FileStream AcquireDirtyMarkerMutationLock(string markerPath)
	{
		string lockPath = markerPath + ".lock";
		DateTime deadlineUtc = DateTime.UtcNow.AddSeconds(5.0);
		while (true)
		{
			try
			{
				return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.WriteThrough);
			}
			catch (IOException)
			{
				if (DateTime.UtcNow >= deadlineUtc)
				{
					throw;
				}
				System.Threading.Thread.Sleep(50);
			}
		}
	}

	private static ProjectTrackingDirtyMarker LoadMarkerCore(string path)
	{
		if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
		{
			return null;
		}
		return DataContractJsonFileStore.Load<ProjectTrackingDirtyMarker>(path);
	}

	private static void WriteMarkerAtomic(string path, ProjectTrackingDirtyMarker marker)
	{
		string temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
		try
		{
			byte[] payload = new UTF8Encoding(false).GetBytes(PlainJsonReportWriter.Serialize(marker));
			using (FileStream stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
			{
				stream.Write(payload, 0, payload.Length);
				stream.Flush(true);
			}
			FamilyBrowserAtomicFileService.Promote(temporaryPath, path);
		}
		finally
		{
			try
			{
				if (File.Exists(temporaryPath))
				{
					File.Delete(temporaryPath);
				}
			}
			catch
			{
			}
		}
	}

	private static bool DirtyMarkerIdentityMatches(ProjectTrackingDirtyMarker current, ProjectTrackingDirtyMarker expected)
	{
		if (current == null || expected == null)
		{
			return false;
		}
		if (!string.IsNullOrWhiteSpace(current.MarkerId) || !string.IsNullOrWhiteSpace(expected.MarkerId))
		{
			return !string.IsNullOrWhiteSpace(current.MarkerId)
				&& !string.IsNullOrWhiteSpace(expected.MarkerId)
				&& string.Equals(current.MarkerId, expected.MarkerId, StringComparison.OrdinalIgnoreCase);
		}
		return string.Equals(current.DetectedAtUtc ?? string.Empty, expected.DetectedAtUtc ?? string.Empty, StringComparison.Ordinal)
			&& string.Equals(current.User ?? string.Empty, expected.User ?? string.Empty, StringComparison.OrdinalIgnoreCase)
			&& string.Equals(current.Reason ?? string.Empty, expected.Reason ?? string.Empty, StringComparison.Ordinal);
	}

	private static List<ProjectTrackingDirtyItem> MergeDirtyMarkerItems(IEnumerable<ProjectTrackingDirtyItem> existing, IEnumerable<ProjectTrackingDirtyItem> incoming)
	{
		return (existing ?? Enumerable.Empty<ProjectTrackingDirtyItem>())
			.Concat(incoming ?? Enumerable.Empty<ProjectTrackingDirtyItem>())
			.Where(x => x != null)
			.GroupBy(BuildDirtyItemIdentity, StringComparer.OrdinalIgnoreCase)
			.Select(x => x.Last())
			.Take(1000)
			.ToList();
	}

	private static string BuildDirtyItemIdentity(ProjectTrackingDirtyItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return string.Join("|", new[]
		{
			item.Action ?? string.Empty,
			item.Kind ?? string.Empty,
			item.Name ?? string.Empty,
			item.CategoryName ?? string.Empty,
			item.ElementIdText ?? string.Empty,
			item.State ?? string.Empty,
			item.RequiredAction ?? string.Empty
		});
	}

	private static string MergeDirtyMarkerReasons(string existing, string incoming)
	{
		return string.Join(" | ", new[] { existing, incoming }
			.Where(x => !string.IsNullOrWhiteSpace(x))
			.Select(x => x.Trim())
			.Distinct(StringComparer.OrdinalIgnoreCase));
	}

	private static Schema EnsureSchema()
	{
		Schema trackingSchema = Schema.Lookup(TrackingSchemaGuid);
		if (trackingSchema != null)
		{
			return trackingSchema;
		}
		SchemaBuilder schemaBuilder = new SchemaBuilder(TrackingSchemaGuid);
		schemaBuilder.SetSchemaName("KKYFamilyBrowserProjectTracking");
		schemaBuilder.SetReadAccessLevel(AccessLevel.Public);
		schemaBuilder.SetWriteAccessLevel(AccessLevel.Public);
		schemaBuilder.SetDocumentation("KKY Family Browser tracked project state catalog.");
		schemaBuilder.AddSimpleField("CatalogJson", typeof(string));
		return schemaBuilder.Finish();
	}

	private static DataStorage FindStorage(Document doc, Schema schema)
	{
		foreach (Element item in new FilteredElementCollector(doc).OfClass(typeof(DataStorage)))
		{
			if (item is DataStorage storage && storage.GetEntity(schema).IsValid())
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
