using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

public sealed class FamilyBrowserRequestAttachmentCopyResult
{
	public string StoredPath { get; set; }

	public string ContentSha256 { get; set; }

	public long SizeBytes { get; set; }

	public bool Created { get; set; }

	public FamilyBrowserRequestAttachmentCopyResult()
	{
		StoredPath = string.Empty;
		ContentSha256 = string.Empty;
		SizeBytes = 0L;
		Created = false;
	}
}

public static class FamilyBrowserRequestFileTransactionService
{
	private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(false);

	public static FamilyBrowserRequestAttachmentCopyResult CopyContentAddressed(string sourcePath, string destinationFolder, string displayName)
	{
		if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
		{
			throw new FileNotFoundException("Request attachment was not found.", sourcePath);
		}
		if (string.IsNullOrWhiteSpace(destinationFolder))
		{
			throw new ArgumentException("Request attachment destination is empty.", "destinationFolder");
		}

		Directory.CreateDirectory(destinationFolder);
		string temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(Path.Combine(destinationFolder, "attachment.bin"));
		try
		{
			long sizeBytes;
			string contentHash = CopyAndHash(sourcePath, temporaryPath, out sizeBytes);
			string storedPath = Path.Combine(destinationFolder, BuildStoredFileName(displayName, contentHash));
			if (File.Exists(storedPath))
			{
				if (!FileMatches(storedPath, sizeBytes, contentHash))
				{
					throw new IOException("An attachment hash path already exists with different content: " + storedPath);
				}
				TryDelete(temporaryPath);
				return new FamilyBrowserRequestAttachmentCopyResult
				{
					StoredPath = storedPath,
					ContentSha256 = contentHash,
					SizeBytes = sizeBytes,
					Created = false
				};
			}

			try
			{
				File.Move(temporaryPath, storedPath);
			}
			catch (IOException)
			{
				if (!File.Exists(storedPath) || !FileMatches(storedPath, sizeBytes, contentHash))
				{
					throw;
				}
				TryDelete(temporaryPath);
				return new FamilyBrowserRequestAttachmentCopyResult
				{
					StoredPath = storedPath,
					ContentSha256 = contentHash,
					SizeBytes = sizeBytes,
					Created = false
				};
			}

			return new FamilyBrowserRequestAttachmentCopyResult
			{
				StoredPath = storedPath,
				ContentSha256 = contentHash,
				SizeBytes = sizeBytes,
				Created = true
			};
		}
		finally
		{
			TryDelete(temporaryPath);
		}
	}

	public static void WriteImmutableText(string path, string contents)
	{
		if (string.IsNullOrWhiteSpace(path))
		{
			throw new ArgumentException("Immutable audit path is empty.", "path");
		}
		string folder = Path.GetDirectoryName(Path.GetFullPath(path));
		if (string.IsNullOrWhiteSpace(folder))
		{
			throw new IOException("Immutable audit folder could not be resolved.");
		}
		Directory.CreateDirectory(folder);
		if (File.Exists(path))
		{
			throw new IOException("Immutable audit entry already exists: " + path);
		}

		string temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
		try
		{
			using (FileStream stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
			using (StreamWriter writer = new StreamWriter(stream, Utf8NoBom))
			{
				writer.Write(contents ?? string.Empty);
				writer.Flush();
				stream.Flush(true);
			}
			File.Move(temporaryPath, path);
		}
		finally
		{
			TryDelete(temporaryPath);
		}
	}

	public static void RollbackCreatedFile(string path)
	{
		TryDelete(path);
	}

	private static string CopyAndHash(string sourcePath, string temporaryPath, out long sizeBytes)
	{
		byte[] buffer = new byte[131072];
		long total = 0L;
		using (SHA256 hash = SHA256.Create())
		using (FileStream source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read))
		using (FileStream destination = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
		{
			int read;
			while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
			{
				hash.TransformBlock(buffer, 0, read, null, 0);
				destination.Write(buffer, 0, read);
				total = checked(total + read);
			}
			hash.TransformFinalBlock(new byte[0], 0, 0);
			destination.Flush(true);
			sizeBytes = total;
			return ToHex(hash.Hash);
		}
	}

	private static bool FileMatches(string path, long expectedSize, string expectedHash)
	{
		FileInfo info = new FileInfo(path);
		if (!info.Exists || info.Length != expectedSize)
		{
			return false;
		}
		using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
		using (SHA256 hash = SHA256.Create())
		{
			return string.Equals(ToHex(hash.ComputeHash(stream)), expectedHash, StringComparison.OrdinalIgnoreCase);
		}
	}

	private static string BuildStoredFileName(string displayName, string contentHash)
	{
		string sourceName = Path.GetFileName(string.IsNullOrWhiteSpace(displayName) ? "attachment" : displayName.Trim());
		string extension = Path.GetExtension(sourceName);
		if (extension.Length > 12)
		{
			extension = extension.Substring(0, 12);
		}
		string baseName = MakeSafePart(Path.GetFileNameWithoutExtension(sourceName));
		if (baseName.Length > 24)
		{
			baseName = baseName.Substring(0, 24);
		}
		if (baseName.Length == 0)
		{
			baseName = "attachment";
		}
		string hashPart = (contentHash ?? string.Empty).ToUpperInvariant();
		if (hashPart.Length > 24)
		{
			hashPart = hashPart.Substring(0, 24);
		}
		return baseName + "-" + hashPart + extension;
	}

	private static string MakeSafePart(string value)
	{
		StringBuilder builder = new StringBuilder();
		char[] invalid = Path.GetInvalidFileNameChars();
		foreach (char ch in value ?? string.Empty)
		{
			builder.Append(Array.IndexOf(invalid, ch) >= 0 ? '_' : ch);
		}
		return builder.ToString().Trim().TrimEnd('.');
	}

	private static string ToHex(byte[] value)
	{
		StringBuilder builder = new StringBuilder(value == null ? 0 : value.Length * 2);
		if (value != null)
		{
			for (int index = 0; index < value.Length; index++)
			{
				builder.Append(value[index].ToString("X2", CultureInfo.InvariantCulture));
			}
		}
		return builder.ToString();
	}

	private static void TryDelete(string path)
	{
		try
		{
			if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
			{
				File.Delete(path);
			}
		}
		catch
		{
		}
	}
}
