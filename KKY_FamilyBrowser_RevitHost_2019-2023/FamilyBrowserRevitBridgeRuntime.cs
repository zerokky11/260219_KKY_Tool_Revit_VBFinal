using System.Runtime.CompilerServices;
using System.Threading;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserRevitBridgeRuntime
{
	private static readonly object SyncRoot = RuntimeHelpers.GetObjectValue(new object());

	private static FamilyBrowserRevitBridgeExternalEventHandler _handler;

	private static ExternalEvent _externalEvent;

	private static FamilyBrowserRevitBridgeServer _server;

	private FamilyBrowserRevitBridgeRuntime()
	{
	}

	public static void Start()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (_server == null)
			{
				_handler = new FamilyBrowserRevitBridgeExternalEventHandler();
				_externalEvent = ExternalEvent.Create((IExternalEventHandler)(object)_handler);
				_server = new FamilyBrowserRevitBridgeServer(_handler, _externalEvent);
				_server.Start();
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static void Stop()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (_server != null)
			{
				_server.Stop();
				_server = null;
			}
			_externalEvent = null;
			_handler = null;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}
}
