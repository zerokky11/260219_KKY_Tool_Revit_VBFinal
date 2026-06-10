using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

public sealed class SystemTypeFingerprintService
{
	private SystemTypeFingerprintService()
	{
	}

	public static string Compute(SystemTypeSemanticSnapshot snapshot)
	{
		if (snapshot == null)
		{
			throw new ArgumentNullException("snapshot");
		}
		List<string> lines = new List<string>
		{
			"SYSFP|v3",
			"S1|" + Normalize(snapshot.SystemFamilyKind),
			"S2|" + Normalize(snapshot.CategoryName),
			"S3|" + Normalize(snapshot.TypeName),
			"S4|" + Normalize(snapshot.ClassificationCode),
			"S5|" + Normalize(snapshot.SegmentName),
			"S6|" + Normalize(snapshot.MaterialName),
			"S7|" + Normalize(snapshot.Shape),
			"S8|" + ProjectSnapshotFingerprintService.NormalizeRoutingPreferenceSignature(snapshot.RoutingPreferenceSignature),
			"S9|" + Normalize(snapshot.CompoundStructureSignature)
		};
		string canonical = string.Join("\n", lines);
		return "sha256:" + ComputeSha256(canonical);
	}

	public static string ComputeSimpleTypeFingerprint(string libraryFamilyId, string typeName)
	{
		return "sha256:" + ComputeSha256("T|" + Normalize(libraryFamilyId) + "|" + Normalize(typeName));
	}

	private static string Normalize(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().Replace("\r\n", "\n").Replace("\r", "\n")
			.ToLowerInvariant();
	}

	private static string ComputeSha256(string text)
	{
		byte[] bytes = Encoding.UTF8.GetBytes(text ?? string.Empty);
		using SHA256 sha = SHA256.Create();
		byte[] hash = sha.ComputeHash(bytes);
		StringBuilder sb = new StringBuilder(checked(hash.Length * 2));
		byte[] array = hash;
		foreach (byte b in array)
		{
			sb.Append(b.ToString("x2", CultureInfo.InvariantCulture));
		}
		return sb.ToString();
	}
}
