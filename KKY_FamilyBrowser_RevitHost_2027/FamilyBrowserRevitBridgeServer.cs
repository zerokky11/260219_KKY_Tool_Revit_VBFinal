using System;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserRevitBridgeServer
{
	private readonly FamilyBrowserRevitBridgeExternalEventHandler _handler;

	private readonly ExternalEvent _externalEvent;

	private readonly ManualResetEvent _stopRequested;

	private Thread _thread;

	public FamilyBrowserRevitBridgeServer(FamilyBrowserRevitBridgeExternalEventHandler handler, ExternalEvent externalEvent)
	{
		_stopRequested = new ManualResetEvent(initialState: false);
		if (handler == null)
		{
			throw new ArgumentNullException("handler");
		}
		if (externalEvent == null)
		{
			throw new ArgumentNullException("externalEvent");
		}
		_handler = handler;
		_externalEvent = externalEvent;
	}

	public void Start()
	{
		if (_thread == null)
		{
			_thread = new Thread(RunLoop);
			_thread.IsBackground = true;
			_thread.Name = "KKY Family Browser Revit Bridge";
			_thread.Start();
		}
	}

	public void Stop()
	{
		_stopRequested.Set();
		TryWakeServer();
		if (_thread != null && _thread.IsAlive)
		{
			_thread.Join(500);
		}
		_thread = null;
	}

	private void RunLoop()
	{
		while (!_stopRequested.WaitOne(0))
		{
			try
			{
				using NamedPipeServerStream pipe = new NamedPipeServerStream("KKY_FamilyBrowser_RevitBridge", PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.None);
				pipe.WaitForConnection();
				if (_stopRequested.WaitOne(0))
				{
					break;
				}
				FamilyBrowserBridgeResponse response;
				using (StreamReader reader = new StreamReader(pipe, Encoding.UTF8, detectEncodingFromByteOrderMarks: false, 4096, leaveOpen: true))
				{
					string requestJson = reader.ReadLine();
					response = HandleRequestJson(requestJson);
				}
				using StreamWriter writer = new StreamWriter(pipe, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), 4096, leaveOpen: true);
				writer.AutoFlush = true;
				writer.WriteLine(PlainJsonReportWriter.Serialize(response));
			}
			catch (IOException ex)
			{
				ProjectData.SetProjectError(ex);
				IOException ex2 = ex;
				if (_stopRequested.WaitOne(0))
				{
					ProjectData.ClearProjectError();
					break;
				}
				ProjectData.ClearProjectError();
			}
			catch (ObjectDisposedException ex3)
			{
				ProjectData.SetProjectError(ex3);
				ObjectDisposedException ex4 = ex3;
				if (_stopRequested.WaitOne(0))
				{
					ProjectData.ClearProjectError();
					break;
				}
				ProjectData.ClearProjectError();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				Thread.Sleep(100);
				ProjectData.ClearProjectError();
			}
		}
	}

	private FamilyBrowserBridgeResponse HandleRequestJson(string requestJson)
	{
		FamilyBrowserBridgeRequest request = null;
		FamilyBrowserBridgeResponse HandleRequestJson;
		try
		{
			if (string.IsNullOrWhiteSpace(requestJson))
			{
				HandleRequestJson = new FamilyBrowserBridgeResponse
				{
					Success = false,
					Message = "Empty bridge request."
				};
			}
			else
			{
				request = DataContractJsonTextStore.Load<FamilyBrowserBridgeRequest>(requestJson);
				int timeoutMilliseconds = ResolveTimeoutMilliseconds(request);
				HandleRequestJson = _handler.ExecuteRequest(request, _externalEvent, timeoutMilliseconds);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			HandleRequestJson = new FamilyBrowserBridgeResponse
			{
				RequestId = ((request == null) ? string.Empty : request.RequestId),
				Success = false,
				Message = ex2.Message
			};
			ProjectData.ClearProjectError();
		}
		return HandleRequestJson;
	}

	private static int ResolveTimeoutMilliseconds(FamilyBrowserBridgeRequest request)
	{
		if (request == null)
		{
			return 5000;
		}
		string left = Normalize(request.Command);
		if (Operators.CompareString(left, Normalize("ApplyStandardFamilies"), TextCompare: false) == 0 || Operators.CompareString(left, Normalize("ApplySystemTypes"), TextCompare: false) == 0)
		{
			return 600000;
		}
		if (Operators.CompareString(left, Normalize("CheckCurrentModel"), TextCompare: false) == 0 || Operators.CompareString(left, Normalize("RunSystemPreflight"), TextCompare: false) == 0)
		{
			return 180000;
		}
		return 15000;
	}

	private static void TryWakeServer()
	{
		try
		{
			using NamedPipeClientStream client = new NamedPipeClientStream(".", "KKY_FamilyBrowser_RevitBridge", PipeDirection.Out);
			client.Connect(100);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static string Normalize(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().ToLowerInvariant();
	}
}
