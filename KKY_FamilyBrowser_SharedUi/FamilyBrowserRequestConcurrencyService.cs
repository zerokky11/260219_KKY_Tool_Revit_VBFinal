using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

public sealed class FamilyBrowserRequestConflictException : InvalidOperationException
{
	public string RequestId { get; private set; }

	public long ExpectedRevision { get; private set; }

	public long CurrentRevision { get; private set; }

	public bool LockTimedOut { get; private set; }

	public FamilyBrowserRequestConflictException(string requestId, long expectedRevision, long currentRevision, bool lockTimedOut)
		: base(BuildMessage(requestId, expectedRevision, currentRevision, lockTimedOut))
	{
		RequestId = requestId ?? string.Empty;
		ExpectedRevision = expectedRevision;
		CurrentRevision = currentRevision;
		LockTimedOut = lockTimedOut;
	}

	private static string BuildMessage(string requestId, long expectedRevision, long currentRevision, bool lockTimedOut)
	{
		if (lockTimedOut)
		{
			return "The request is being changed by another user. RequestId=" + (requestId ?? string.Empty);
		}
		return "The request changed after this screen was rendered. RequestId=" + (requestId ?? string.Empty)
			+ ", expectedRevision=" + expectedRevision.ToString(CultureInfo.InvariantCulture)
			+ ", currentRevision=" + currentRevision.ToString(CultureInfo.InvariantCulture);
	}
}

public sealed class FamilyBrowserRequestMutationLease : IDisposable
{
	private FileStream _stream;

	internal FamilyBrowserRequestMutationLease(FileStream stream)
	{
		_stream = stream;
	}

	public void Dispose()
	{
		FileStream stream = Interlocked.Exchange(ref _stream, null);
		if (stream != null)
		{
			stream.Dispose();
		}
	}
}

public static class FamilyBrowserRequestConcurrencyService
{
	public static FamilyBrowserRequestMutationLease Acquire(string requestStoreFolder, string requestId, int timeoutMilliseconds = 2500)
	{
		if (string.IsNullOrWhiteSpace(requestStoreFolder))
		{
			throw new ArgumentException("Request store folder is empty.", "requestStoreFolder");
		}
		if (string.IsNullOrWhiteSpace(requestId))
		{
			throw new ArgumentException("Request id is empty.", "requestId");
		}

		Directory.CreateDirectory(requestStoreFolder);
		string lockPath = Path.Combine(requestStoreFolder, ".kky-r-" + StableKey(requestId).Substring(0, 16) + ".lck");
		DateTime deadline = DateTime.UtcNow.AddMilliseconds(Math.Max(100, timeoutMilliseconds));
		while (true)
		{
			try
			{
				FileStream stream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.WriteThrough);
				return new FamilyBrowserRequestMutationLease(stream);
			}
			catch (IOException)
			{
				if (DateTime.UtcNow >= deadline)
				{
					throw new FamilyBrowserRequestConflictException(requestId, 0L, 0L, true);
				}
				Thread.Sleep(75);
			}
		}
	}

	public static void EnsureExpectedRevision(string requestId, long expectedRevision, string expectedToken, long currentRevision, string currentToken)
	{
		bool revisionMatches = expectedRevision > 0L && currentRevision > 0L && expectedRevision == currentRevision;
		bool tokenMatches = !string.IsNullOrWhiteSpace(expectedToken)
			&& !string.IsNullOrWhiteSpace(currentToken)
			&& string.Equals(expectedToken.Trim(), currentToken.Trim(), StringComparison.OrdinalIgnoreCase);
		if (!revisionMatches || !tokenMatches)
		{
			throw new FamilyBrowserRequestConflictException(requestId, expectedRevision, currentRevision, false);
		}
	}

	public static string CreateRevisionToken()
	{
		return Guid.NewGuid().ToString("N");
	}

	public static string ComputeFileToken(string path)
	{
		if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
		{
			return string.Empty;
		}
		using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
		using (SHA256 hash = SHA256.Create())
		{
			return "legacy-" + ToHex(hash.ComputeHash(stream));
		}
	}

	private static string StableKey(string value)
	{
		using (SHA256 hash = SHA256.Create())
		{
			return ToHex(hash.ComputeHash(Encoding.UTF8.GetBytes((value ?? string.Empty).Trim().ToUpperInvariant())));
		}
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
}
