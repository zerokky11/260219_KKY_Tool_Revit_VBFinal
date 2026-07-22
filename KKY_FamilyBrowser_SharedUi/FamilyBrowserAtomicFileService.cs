using System;
using System.IO;

public static class FamilyBrowserAtomicFileService
{
	public static string CreateSiblingTemporaryPath(string destinationPath)
	{
		return CreateSiblingAuxiliaryPath(destinationPath, ".kky-t-", ".tmp");
	}

	public static string CreateSiblingBackupPath(string destinationPath)
	{
		return CreateSiblingAuxiliaryPath(destinationPath, ".kky-b-", ".bak");
	}

	public static void Promote(string temporaryPath, string destinationPath)
	{
		if (string.IsNullOrWhiteSpace(temporaryPath))
		{
			throw new ArgumentException("Temporary path is empty.", "temporaryPath");
		}
		if (string.IsNullOrWhiteSpace(destinationPath))
		{
			throw new ArgumentException("Destination path is empty.", "destinationPath");
		}
		if (!File.Exists(destinationPath))
		{
			File.Move(temporaryPath, destinationPath);
			return;
		}

		try
		{
			File.Replace(temporaryPath, destinationPath, null, true);
			return;
		}
		catch
		{
			// Some network shares reject Replace. Keep the committed file recoverable until promotion succeeds.
		}

		string backupPath = CreateSiblingBackupPath(destinationPath);
		File.Move(destinationPath, backupPath);
		try
		{
			File.Move(temporaryPath, destinationPath);
			TryDelete(backupPath);
		}
		catch
		{
			if (!File.Exists(destinationPath) && File.Exists(backupPath))
			{
				File.Move(backupPath, destinationPath);
			}
			throw;
		}
	}

	private static string CreateSiblingAuxiliaryPath(string destinationPath, string prefix, string extension)
	{
		if (string.IsNullOrWhiteSpace(destinationPath))
		{
			throw new ArgumentException("Destination path is empty.", "destinationPath");
		}
		string directory = Path.GetDirectoryName(Path.GetFullPath(destinationPath));
		if (string.IsNullOrWhiteSpace(directory))
		{
			throw new IOException("The destination folder could not be resolved.");
		}
		for (int attempt = 0; attempt < 16; attempt++)
		{
			string candidate = Path.Combine(directory, prefix + Guid.NewGuid().ToString("N").Substring(0, 8) + extension);
			if (!File.Exists(candidate) && !Directory.Exists(candidate))
			{
				return candidate;
			}
		}
		throw new IOException("A unique sibling file path could not be allocated.");
	}

	private static void TryDelete(string path)
	{
		try
		{
			if (File.Exists(path))
			{
				File.Delete(path);
			}
		}
		catch
		{
		}
	}
}
