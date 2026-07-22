using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyThumbnailConstraintDialogGuard : IDisposable
{
	private enum NativeFamilyScanDialogAction
	{
		None,
		Confirm,
		Cancel
	}

	private sealed class NativeFamilyScanDialog
	{
		public IntPtr Handle { get; }

		public string Title { get; }

		public string Body { get; }

		public NativeFamilyScanDialogAction Action { get; }

		public string Reason { get; }

		public List<NativeWindowButton> Buttons { get; }

		public NativeFamilyScanDialog(IntPtr handle, string title, string body, NativeFamilyScanDialogAction action, string reason, IEnumerable<NativeWindowButton> buttons = null)
		{
			Handle = handle;
			Title = title ?? string.Empty;
			Body = body ?? string.Empty;
			Action = action;
			Reason = reason ?? string.Empty;
			Buttons = (buttons ?? Enumerable.Empty<NativeWindowButton>()).Where([SpecialName] (NativeWindowButton x) => x != null).ToList();
		}
	}

	private sealed class NativeWindowButton
	{
		public IntPtr Handle { get; }

		public string Text { get; }

		public bool Enabled { get; }

		public int ControlId { get; }

		public NativeWindowButton(IntPtr handle, string text, bool enabled, int controlId)
		{
			Handle = handle;
			Text = text ?? string.Empty;
			Enabled = enabled;
			ControlId = controlId;
		}
	}

	private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

	private readonly UIApplication _uiApplication;

	private bool _disposed;

	private readonly List<FamilyThumbnailAutoConfirmedDialogRecord> _records;

	private readonly object _recordSyncRoot;

	private readonly Timer _nativeDialogTimer;

	private int _nativeDialogTicking;

	private IntPtr _lastNativeDialogHandle;

	private DateTime _lastNativeDialogHandledAtUtc;

	private string _currentFamilyName;

	private string _currentCategoryName;

	private const int BM_CLICK = 245;

	private const int WM_CLOSE = 16;

	private const int WM_COMMAND = 273;

	private const int IDOK = 1;

	private const int IDCANCEL = 2;

	public int HandledConstraintWarnings => RecordCount;

	public int RecordCount
	{
		get
		{
			object recordSyncRoot = _recordSyncRoot;
			ObjectFlowControl.CheckForSyncLockOnValueType(recordSyncRoot);
			bool lockTaken = false;
			try
			{
				Monitor.Enter(recordSyncRoot, ref lockTaken);
				return _records.Count;
			}
			finally
			{
				if (lockTaken)
				{
					Monitor.Exit(recordSyncRoot);
				}
			}
		}
	}

	public FamilyThumbnailConstraintDialogGuard(UIApplication uiApplication)
	{
		_records = new List<FamilyThumbnailAutoConfirmedDialogRecord>();
		_recordSyncRoot = RuntimeHelpers.GetObjectValue(new object());
		_lastNativeDialogHandle = IntPtr.Zero;
		_lastNativeDialogHandledAtUtc = DateTime.MinValue;
		_currentFamilyName = string.Empty;
		_currentCategoryName = string.Empty;
		_uiApplication = uiApplication;
		if (_uiApplication != null)
		{
			_uiApplication.DialogBoxShowing += HandleDialogBoxShowing;
		}
		_nativeDialogTimer = new Timer(HandleNativeDialogTimerTick, null, 150, 150);
	}

	public void SetCurrentFamily(string categoryName, string familyName)
	{
		lock (_recordSyncRoot)
		{
			_currentCategoryName = categoryName ?? string.Empty;
			_currentFamilyName = familyName ?? string.Empty;
		}
	}

	public void ClearCurrentFamily()
	{
		lock (_recordSyncRoot)
		{
			_currentCategoryName = string.Empty;
			_currentFamilyName = string.Empty;
		}
	}

	private bool TryGetCurrentFamilyContext(out string categoryName, out string familyName)
	{
		lock (_recordSyncRoot)
		{
			categoryName = _currentCategoryName ?? string.Empty;
			familyName = _currentFamilyName ?? string.Empty;
		}
		return !string.IsNullOrWhiteSpace(familyName);
	}

	public List<FamilyThumbnailAutoConfirmedDialogRecord> GetRecordsSince(int startIndex)
	{
		if (startIndex < 0)
		{
			startIndex = 0;
		}
		object recordSyncRoot = _recordSyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(recordSyncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(recordSyncRoot, ref lockTaken);
			if (startIndex >= _records.Count)
			{
				return new List<FamilyThumbnailAutoConfirmedDialogRecord>();
			}
			return _records.Skip(startIndex).ToList();
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(recordSyncRoot);
			}
		}
	}

	private void HandleDialogBoxShowing(object sender, DialogBoxShowingEventArgs e)
	{
		if (e == null)
		{
			return;
		}
		string currentCategoryName;
		string currentFamilyName;
		bool activeFamilyEditScope = TryGetCurrentFamilyContext(out currentCategoryName, out currentFamilyName);
		string dialogText = CollectDialogText(e);
		string reason = ResolveAutoConfirmReason(dialogText);
		if (string.IsNullOrWhiteSpace(reason) && activeFamilyEditScope && IsOpeningNotCuttingAnythingText(dialogText))
		{
			reason = "OpeningNotCuttingAnything";
		}
		if (ShouldCancelDialog(reason))
		{
			// Delete/Remove dialogs must be decided from their real enabled buttons by the native timer.
			return;
		}
		if (!string.IsNullOrWhiteSpace(reason))
		{
			string resultText = string.Empty;
			resultText = TryOverrideDialogResult(e, new int[3] { 1, 1, 8 }, new string[3] { "IDOK(1)", "TaskDialogResult.Ok", "TaskDialogResult.Close" });
			if (!string.IsNullOrWhiteSpace(resultText))
			{
				AddRecord(new FamilyThumbnailAutoConfirmedDialogRecord
				{
					ConfirmedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
					CategoryName = currentCategoryName,
					FamilyName = currentFamilyName,
					Reason = reason,
					DialogText = dialogText,
					OverrideResult = resultText,
					ActionTaken = "OK",
					AvailableButtons = "DialogBoxShowingEventArgs"
				});
			}
		}
	}

	private void AddRecord(FamilyThumbnailAutoConfirmedDialogRecord record)
	{
		if (record == null)
		{
			return;
		}
		object recordSyncRoot = _recordSyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(recordSyncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(recordSyncRoot, ref lockTaken);
			_records.Add(record);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(recordSyncRoot);
			}
		}
	}

	private void HandleNativeDialogTimerTick(object state)
	{
		if (_disposed || Interlocked.Exchange(ref _nativeDialogTicking, 1) == 1)
		{
			return;
		}
		try
		{
			string currentCategoryName;
			string currentFamilyName;
			bool activeFamilyEditScope = TryGetCurrentFamilyContext(out currentCategoryName, out currentFamilyName);
			NativeFamilyScanDialog candidate = FindNativeFamilyScanDialog(activeFamilyEditScope);
			if (candidate != null && !(candidate.Handle == IntPtr.Zero) && (!(candidate.Handle == _lastNativeDialogHandle) || !((DateTime.UtcNow - _lastNativeDialogHandledAtUtc).TotalSeconds < 3.0)))
			{
				string clickedButtonText = string.Empty;
				if (TryClickNativeDialogAction(candidate, ref clickedButtonText))
				{
					_lastNativeDialogHandle = candidate.Handle;
					_lastNativeDialogHandledAtUtc = DateTime.UtcNow;
					AddRecord(new FamilyThumbnailAutoConfirmedDialogRecord
					{
						ConfirmedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
						CategoryName = currentCategoryName,
						FamilyName = currentFamilyName,
						Reason = "NativeDialog:" + candidate.Reason,
						DialogText = candidate.Title + Environment.NewLine + candidate.Body,
						OverrideResult = "Win32Click:" + clickedButtonText,
						ActionTaken = ((candidate.Action == NativeFamilyScanDialogAction.Cancel) ? "Cancel" : "OK"),
						AvailableButtons = BuildNativeButtonSummary(candidate.Buttons)
					});
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		finally
		{
			Interlocked.Exchange(ref _nativeDialogTicking, 0);
		}
	}

	private static string CollectDialogText(DialogBoxShowingEventArgs e)
	{
		List<string> values = new List<string>();
		HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
		AddDialogTextValue(values, seen, "EventType", e.GetType().FullName);
		try
		{
			string summary = e.ToString();
			if (!string.IsNullOrWhiteSpace(summary))
			{
				AddDialogTextValue(values, seen, "ToString", summary);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		AddDialogMemberValues(values, seen, e.GetType(), e, BindingFlags.Instance | BindingFlags.Public);
		AddDialogMemberValues(values, seen, e.GetType(), e, BindingFlags.Instance | BindingFlags.NonPublic);
		return string.Join(Environment.NewLine, values);
	}

	private static void AddDialogMemberValues(IList<string> values, ISet<string> seen, Type sourceType, object instance, BindingFlags flags)
	{
		if (values == null || seen == null || (object)sourceType == null || instance == null)
		{
			return;
		}
		PropertyInfo[] properties = sourceType.GetProperties(flags);
		foreach (PropertyInfo propertyInfo in properties)
		{
			if ((object)propertyInfo != null && propertyInfo.GetIndexParameters().Length <= 0)
			{
				try
				{
					object value = RuntimeHelpers.GetObjectValue(propertyInfo.GetValue(RuntimeHelpers.GetObjectValue(instance), null));
					AddDialogTextValue(values, seen, propertyInfo.Name, RuntimeHelpers.GetObjectValue(value));
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
		FieldInfo[] fields = sourceType.GetFields(flags);
		foreach (FieldInfo fieldInfo in fields)
		{
			if ((object)fieldInfo != null)
			{
				try
				{
					object value2 = RuntimeHelpers.GetObjectValue(fieldInfo.GetValue(RuntimeHelpers.GetObjectValue(instance)));
					AddDialogTextValue(values, seen, fieldInfo.Name, RuntimeHelpers.GetObjectValue(value2));
				}
				catch (Exception projectError2)
				{
					ProjectData.SetProjectError(projectError2);
					ProjectData.ClearProjectError();
				}
			}
		}
	}

	private static void AddDialogTextValue(IList<string> values, ISet<string> seen, string label, object value)
	{
		if (values == null || seen == null || value == null)
		{
			return;
		}
		Type valueType = value.GetType();
		if ((object)valueType != typeof(string) && !valueType.IsValueType && !valueType.IsEnum)
		{
			return;
		}
		string text;
		try
		{
			text = Convert.ToString(RuntimeHelpers.GetObjectValue(value), CultureInfo.InvariantCulture);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
			return;
		}
		if (!string.IsNullOrWhiteSpace(text))
		{
			if (text.Length > 4000)
			{
				text = text.Substring(0, 4000) + "...";
			}
			string entry = (string.IsNullOrWhiteSpace(label) ? text : (label + "=" + text));
			if (!seen.Contains(entry))
			{
				seen.Add(entry);
				values.Add(entry);
			}
		}
	}

	private static string TryOverrideDialogResult(DialogBoxShowingEventArgs e, int[] resultValues, string[] resultLabels)
	{
		if (e == null || resultValues == null)
		{
			return string.Empty;
		}
		checked
		{
			int num = resultValues.Length - 1;
			for (int index = 0; index <= num; index++)
			{
				try
				{
					e.OverrideResult(resultValues[index]);
					if (resultLabels != null && index < resultLabels.Length)
					{
						return resultLabels[index];
					}
					return resultValues[index].ToString(CultureInfo.InvariantCulture);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
			return string.Empty;
		}
	}

	private static string ResolveAutoConfirmReason(string dialogText)
	{
		if (string.IsNullOrWhiteSpace(dialogText))
		{
			return string.Empty;
		}
		string robustReason = ResolveAutoConfirmReasonFromNormalizedText(dialogText);
		if (!string.IsNullOrWhiteSpace(robustReason))
		{
			return robustReason;
		}
		if (ContainsAll(dialogText, "constraints between geometry in the family", "parameter modification"))
		{
			return "ConstraintParameterWarning";
		}
		if (ContainsAll(dialogText, "constraints between geometry") || ContainsAll(dialogText, "constraint", "geometry", "family") || ContainsAll(dialogText, "constraint", "geometry", "parameter"))
		{
			return "ConstraintParameterWarning";
		}
		if (ContainsAll(dialogText, "remove constraints") || ContainsAll(dialogText, "remove", "constraint"))
		{
			return "RemoveConstraints";
		}
		if (ContainsAll(dialogText, "delete", "constraint") || ContainsAll(dialogText, "구속", "삭제") || ContainsAll(dialogText, "제약", "삭제"))
		{
			return "DeleteConstraints";
		}
		if (ContainsAll(dialogText, "constraints between geometry in the family", "parameter modification"))
		{
			return "ConstraintParameterWarning";
		}
		if (ContainsAll(dialogText, "constraint", "family", "parameter") || ContainsAll(dialogText, "구속", "패밀리", "매개변수"))
		{
			return "FamilyConstraintParameterWarning";
		}
		if (ContainsAll(dialogText, "geometry", "family", "warning") || ContainsAll(dialogText, "geometry", "family", "error") || ContainsAll(dialogText, "geometry", "constraint"))
		{
			return "FamilyGeometryWarning";
		}
		string fallbackReason = ResolveStandardFamilyScanWarningFallback(dialogText);
		if (!string.IsNullOrWhiteSpace(fallbackReason))
		{
			return fallbackReason;
		}
		return string.Empty;
	}

	private static string ResolveAutoConfirmReasonFromNormalizedText(string dialogText)
	{
		string text = NormalizeDialogText(dialogText);
		if (string.IsNullOrWhiteSpace(text))
		{
			return string.Empty;
		}
		if (IsGeometryConstraintWarningText(text))
		{
			return "ConstraintParameterWarning";
		}
		if (IsDestructiveDeleteChoiceText(text))
		{
			return "DeleteInstanceOrType";
		}
		if (ContainsAll(text, "remove constraints") || ContainsAll(text, "delete constraints") || ContainsAll(text, "remove", "constraint") || ContainsAll(text, "delete", "constraint") || ContainsAny(text, "제거", "삭제"))
		{
			if (ContainsAny(text, "delete", "삭제"))
			{
				return "DeleteConstraints";
			}
			return "RemoveConstraints";
		}
		if (ContainsAll(text, "constraints between geometry") || ContainsAll(text, "constraint", "geometry", "family") || ContainsAll(text, "constraint", "geometry", "parameter") || ContainsAll(text, "constraint", "geometry", "warning") || ContainsAll(text, "constraint", "geometry", "invalid") || ContainsAll(text, "constraint", "geometry", "dimension") || ContainsAll(text, "constraints", "geometry") || ContainsAll(text, "warning", "geometry") || ContainsAll(text, "warning", "constraint") || ContainsAll(text, "failed", "geometry") || ContainsAll(text, "cannot", "geometry") || ContainsAll(text, "can not", "geometry"))
		{
			return "ConstraintParameterWarning";
		}
		if (ContainsDialogIdHint(text, "constraint") || ContainsDialogIdHint(text, "constraints") || ContainsDialogIdHint(text, "geometry"))
		{
			if (ContainsAny(text, "remove", "delete", "제거", "삭제"))
			{
				if (ContainsAny(text, "delete", "삭제"))
				{
					return "DeleteConstraints";
				}
				return "RemoveConstraints";
			}
			return "ConstraintParameterWarning";
		}
		if (ContainsAll(text, "remove constraints") || ContainsAll(text, "remove", "constraint") || ContainsAny(text, "removeconstraints", "constraintsremove") || ContainsAll(text, "구속", "제거") || ContainsAll(text, "제약", "제거") || ContainsAll(text, "구속", "제거") || ContainsAll(text, "제약", "제거"))
		{
			return "RemoveConstraints";
		}
		if (ContainsAll(text, "delete constraints") || ContainsAll(text, "delete", "constraint") || ContainsAll(text, "구속", "삭제") || ContainsAll(text, "제약", "삭제") || ContainsAll(text, "구속", "삭제") || ContainsAll(text, "제약", "삭제"))
		{
			return "DeleteConstraints";
		}
		if (ContainsAll(text, "constraints between geometry") || ContainsAll(text, "constraint", "geometry", "family") || ContainsAll(text, "constraint", "geometry", "parameter") || ContainsAll(text, "constraint", "geometry", "warning") || ContainsAll(text, "constraint", "geometry", "invalid") || ContainsAll(text, "constraint", "geometry", "dimension") || ContainsAll(text, "constraints", "geometry") || ContainsAll(text, "warning", "geometry") || ContainsAll(text, "warning", "constraint") || ContainsAll(text, "failed", "geometry") || ContainsAll(text, "cannot", "geometry") || ContainsAll(text, "can not", "geometry") || ContainsAll(text, "형상", "구속") || ContainsAll(text, "기하", "구속") || ContainsAll(text, "형상", "제약") || ContainsAll(text, "패밀리", "구속", "매개변수") || ContainsAll(text, "패밀리", "제약", "매개변수") || ContainsAll(text, "경고", "형상") || ContainsAll(text, "경고", "구속") || ContainsAll(text, "형상", "구속") || ContainsAll(text, "형상", "제약"))
		{
			return "ConstraintParameterWarning";
		}
		if (ContainsAll(text, "constraint", "family", "parameter") || ContainsAll(text, "constraint", "family", "formula") || ContainsAll(text, "constraint", "family", "type") || ContainsAll(text, "패밀리", "구속") || ContainsAll(text, "패밀리", "제약"))
		{
			return "FamilyConstraintParameterWarning";
		}
		if (ContainsAll(text, "geometry", "family", "warning") || ContainsAll(text, "geometry", "family", "error") || ContainsAll(text, "geometry", "constraint") || ContainsAll(text, "형상", "패밀리", "경고") || ContainsAll(text, "기하", "패밀리", "경고"))
		{
			return "FamilyGeometryWarning";
		}
		if (ContainsDialogIdHint(text, "docwarn") || ContainsDialogIdHint(text, "warning") || ContainsDialogIdHint(text, "warn"))
		{
			return "RevitWarning";
		}
		string fallbackReason = ResolveStandardFamilyScanWarningFallback(text);
		if (!string.IsNullOrWhiteSpace(fallbackReason))
		{
			return fallbackReason;
		}
		return string.Empty;
	}

	private static string ResolveStandardFamilyScanWarningFallback(string dialogText)
	{
		string text = NormalizeDialogText(dialogText);
		if (string.IsNullOrWhiteSpace(text))
		{
			return string.Empty;
		}
		if (IsDestructiveDeleteChoiceText(text) || ContainsAny(text, "remove constraints", "delete constraints", "구속 제거", "구속 삭제", "제약 제거", "제약 삭제") || (ContainsAny(text, "remove", "delete", "제거", "삭제") && ContainsAny(text, "constraint", "constraints", "구속", "제약")))
		{
			return string.Empty;
		}
		if (ContainsAny(text, "constraints between geometry", "constraint geometry", "geometry constraint", "형상 구속", "기하 구속", "형상 제약", "기하 제약"))
		{
			return "ConstraintParameterWarning";
		}
		bool looksLikeRevitDialog = ContainsAny(text, "taskdialog", "dialogboxshowingeventargs", "dialogid", "messagebox", "dialog");
		bool looksLikeWarning = ContainsAny(text, "warning", "warn", "docwarn", "경고", "caution", "error", "failed", "cannot", "can not", "오류", "실패", "할 수 없습니다");
		bool looksLikeFamilyScanIssue = ContainsAny(text, "family", "families", "geometry", "constraint", "constraints", "parameter", "formula", "dimension", "thumbnail", "preview", "image", "패밀리", "형상", "기하", "구속", "제약", "매개변수", "치수", "미리보기", "이미지");
		if (looksLikeRevitDialog && looksLikeWarning && looksLikeFamilyScanIssue)
		{
			return "StandardFamilyScanOkDialog";
		}
		return string.Empty;
	}

	private static bool ShouldCancelDialog(string reason)
	{
		if (string.Equals(reason, "RemoveConstraints", StringComparison.OrdinalIgnoreCase) || string.Equals(reason, "DeleteInstanceOrType", StringComparison.OrdinalIgnoreCase))
		{
			return true;
		}
		return string.Equals(reason, "DeleteConstraints", StringComparison.OrdinalIgnoreCase);
	}

	private static bool IsGeometryConstraintWarningText(string text)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return false;
		}
		return ContainsAll(text, "constraints between geometry") || ContainsAll(text, "constraint", "geometry", "family") || ContainsAll(text, "constraint", "geometry", "parameter") || ContainsAll(text, "constraint", "geometry", "warning") || ContainsAll(text, "constraint", "geometry", "invalid") || ContainsAll(text, "constraint", "geometry", "dimension") || ContainsAll(text, "constraints", "geometry") || ContainsAll(text, "warning", "geometry") || ContainsAll(text, "warning", "constraint") || ContainsAll(text, "failed", "geometry") || ContainsAll(text, "cannot", "geometry") || ContainsAll(text, "can not", "geometry");
	}

	private static bool HasRemoveOrDeleteConstraintChoice(string text)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return false;
		}
		return ContainsAny(text, "remove constraints", "delete constraints") || (ContainsAny(text, "remove", "delete") && ContainsAny(text, "constraint", "constraints"));
	}

	private static bool IsDestructiveDeleteChoiceText(string text)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return false;
		}
		return ContainsAny(text, "delete instance", "delete type", "delete instances", "delete types", "delete element", "delete elements", "remove instance", "remove type", "인스턴스 삭제", "타입 삭제", "유형 삭제", "요소 삭제") || (ContainsAny(text, "delete", "remove", "삭제", "제거") && ContainsAny(text, "instance", "instances", "type", "types", "element", "elements", "인스턴스", "타입", "유형", "요소"));
	}

	private static string NormalizeDialogText(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		StringBuilder builder = new StringBuilder(value.Length);
		foreach (char ch in value)
		{
			if (char.IsWhiteSpace(ch) || ch == '_' || ch == '-' || ch == ':' || ch == '.' || ch == ',')
			{
				builder.Append(' ');
			}
			else
			{
				builder.Append(char.ToLowerInvariant(ch));
			}
		}
		return builder.ToString();
	}

	private static bool ContainsAll(string value, params string[] fragments)
	{
		foreach (string fragment in fragments)
		{
			if (value.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) < 0)
			{
				return false;
			}
		}
		return true;
	}

	private static bool ContainsAny(string value, params string[] fragments)
	{
		foreach (string fragment in fragments)
		{
			if (value.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0)
			{
				return true;
			}
		}
		return false;
	}

	private static bool ContainsDialogIdHint(string value, string fragment)
	{
		if (string.IsNullOrWhiteSpace(value) || string.IsNullOrWhiteSpace(fragment))
		{
			return false;
		}
		return ContainsAll(value, "dialogid", fragment) || ContainsAll(value, "taskdialog", fragment);
	}

	private static NativeFamilyScanDialog FindNativeFamilyScanDialog(bool activeFamilyEditScope)
	{
		NativeFamilyScanDialog nativeFamilyScanDialog = null;
		int id = Process.GetCurrentProcess().Id;
		EnumWindows([SpecialName] (IntPtr hWnd, IntPtr lParam) =>
		{
			if (nativeFamilyScanDialog != null)
			{
				return false;
			}
			if (!IsWindowVisible(hWnd))
			{
				return true;
			}
			uint processId = 0u;
			GetWindowThreadProcessId(hWnd, ref processId);
			if (processId != checked((uint)id))
			{
				return true;
			}
			if (!IsPotentialNativeDialogClass(GetWindowClassName(hWnd)))
			{
				return true;
			}
			string windowTextSafe = GetWindowTextSafe(hWnd);
			string body = string.Join(Environment.NewLine, GetDescendantTexts(hWnd));
			List<NativeWindowButton> buttons = GetChildButtons(hWnd);
			NativeFamilyScanDialogAction nativeFamilyScanDialogAction = ResolveNativeFamilyScanDialogAction(windowTextSafe, body, buttons, activeFamilyEditScope);
			if (nativeFamilyScanDialogAction == NativeFamilyScanDialogAction.None)
			{
				return true;
			}
			nativeFamilyScanDialog = new NativeFamilyScanDialog(hWnd, windowTextSafe, body, nativeFamilyScanDialogAction, ResolveNativeFamilyScanDialogReason(nativeFamilyScanDialogAction, windowTextSafe, body), buttons);
			return false;
		}, IntPtr.Zero);
		return nativeFamilyScanDialog;
	}

	public static string BuildFingerprintCanceledReason(IEnumerable<FamilyThumbnailAutoConfirmedDialogRecord> records)
	{
		FamilyThumbnailAutoConfirmedDialogRecord cancelRecord = (records ?? Enumerable.Empty<FamilyThumbnailAutoConfirmedDialogRecord>()).FirstOrDefault([SpecialName] (FamilyThumbnailAutoConfirmedDialogRecord x) => IsFingerprintCancelRecord(x));
		if (cancelRecord == null)
		{
			return string.Empty;
		}
		string reason = (string.IsNullOrWhiteSpace(cancelRecord.Reason) ? "Family edit warning" : cancelRecord.Reason);
		return "Fingerprint was not created because Family Browser automatically pressed Cancel for a protected family edit warning. Reason=" + reason;
	}

	private static bool IsFingerprintCancelRecord(FamilyThumbnailAutoConfirmedDialogRecord record)
	{
		if (record == null)
		{
			return false;
		}
		string resultText = NormalizeDialogText(record.OverrideResult ?? string.Empty);
		if (ContainsAny(resultText, "idok", "taskdialogresult.ok", "ok", "continue", "proceed"))
		{
			return false;
		}
		if (ContainsAny(resultText, "idcancel", "taskdialogresult.cancel", "cancel"))
		{
			return true;
		}
		return ContainsAny(NormalizeDialogText(record.Reason ?? string.Empty), "removeconstraints", "deleteconstraints", "deleteinstanceortype", "remove constraints", "delete constraints", "delete instance", "delete type");
	}

	private static bool IsPotentialNativeDialogClass(string className)
	{
		if (string.IsNullOrWhiteSpace(className))
		{
			return false;
		}
		return string.Equals(className, "#32770", StringComparison.OrdinalIgnoreCase) || className.IndexOf("Dialog", StringComparison.OrdinalIgnoreCase) >= 0 || className.IndexOf("Task", StringComparison.OrdinalIgnoreCase) >= 0;
	}

	private static NativeFamilyScanDialogAction ResolveNativeFamilyScanDialogAction(string title, string body, IList<NativeWindowButton> buttons, bool activeFamilyEditScope)
	{
		string text = NormalizeDialogText((title ?? string.Empty) + "\n" + (body ?? string.Empty));
		if (string.IsNullOrWhiteSpace(text))
		{
			return NativeFamilyScanDialogAction.None;
		}
		bool hasOkButton = (buttons ?? new List<NativeWindowButton>()).Any([SpecialName] (NativeWindowButton x) => x != null && x.Enabled && IsNativeOkButton(x));
		bool hasCancelButton = (buttons ?? new List<NativeWindowButton>()).Any([SpecialName] (NativeWindowButton x) => x != null && x.Enabled && IsNativeCancelButton(x));
		if (activeFamilyEditScope && hasOkButton)
		{
			return NativeFamilyScanDialogAction.Confirm;
		}
		if (activeFamilyEditScope && hasCancelButton && !hasOkButton && HasOnlyDeleteOrCancelButtons(buttons))
		{
			return NativeFamilyScanDialogAction.Cancel;
		}
		if (IsDestructiveDeleteChoiceText(text) || ContainsAny(text, "remove constraints", "delete constraints", "제거", "삭제") || (ContainsAny(text, "remove", "delete") && ContainsAny(text, "constraint", "constraints")))
		{
			if (!hasOkButton && hasCancelButton)
			{
				return NativeFamilyScanDialogAction.Cancel;
			}
		}
		if (hasOkButton && (IsGeometryConstraintWarningText(text) || LooksLikeStandardFamilyScanDialog(text)))
		{
			return NativeFamilyScanDialogAction.Confirm;
		}
		if (ContainsAny(text, "constraints between geometry", "constraint geometry", "geometry constraint") && hasOkButton)
		{
			return NativeFamilyScanDialogAction.Confirm;
		}
		if ((ContainsAny(text, "remove constraints", "delete constraints", "구속 제거", "구속 삭제", "제약 제거", "제약 삭제") || (ContainsAny(text, "remove", "delete", "제거", "삭제") && ContainsAny(text, "constraint", "constraints", "구속", "제약"))) && !hasOkButton && hasCancelButton)
		{
			return NativeFamilyScanDialogAction.Cancel;
		}
		if (ContainsAny(text, "constraints between geometry", "constraint geometry", "geometry constraint", "형상 구속", "기하 구속", "형상 제약", "기하 제약") && hasOkButton)
		{
			return NativeFamilyScanDialogAction.Confirm;
		}
		if (hasCancelButton && !hasOkButton && HasOnlyDeleteOrCancelButtons(buttons))
		{
			return NativeFamilyScanDialogAction.Cancel;
		}
		return NativeFamilyScanDialogAction.None;
	}

	private static string ResolveNativeFamilyScanDialogReason(NativeFamilyScanDialogAction action, string title, string body)
	{
		string text = NormalizeDialogText((title ?? string.Empty) + "\n" + (body ?? string.Empty));
		if (action == NativeFamilyScanDialogAction.Cancel)
		{
			if (IsDestructiveDeleteChoiceText(text))
			{
				return "DeleteInstanceOrType";
			}
			if (ContainsAny(text, "delete", "삭제"))
			{
				return "DeleteConstraints";
			}
			return "RemoveConstraints";
		}
		if (ContainsAny(text, "constraints between geometry", "constraint geometry", "geometry constraint", "형상 구속", "기하 구속", "형상 제약", "기하 제약"))
		{
			return "ConstraintParameterWarning";
		}
		if (IsOpeningNotCuttingAnythingText(text))
		{
			return "OpeningNotCuttingAnything";
		}
		return "StandardFamilyScanOkDialog";
	}

	public static string ResolveFamilyEditDialogActionForAudit(string title, string body, IEnumerable<string> enabledButtonLabels, bool activeFamilyEditScope)
	{
		List<NativeWindowButton> buttons = new List<NativeWindowButton>();
		int nextControlId = 100;
		foreach (string rawLabel in enabledButtonLabels ?? Enumerable.Empty<string>())
		{
			string label = rawLabel ?? string.Empty;
			bool enabled = true;
			int explicitControlId = 0;
			if (label.StartsWith("disabled:", StringComparison.OrdinalIgnoreCase))
			{
				enabled = false;
				label = label.Substring("disabled:".Length);
			}
			if (label.StartsWith("idok:", StringComparison.OrdinalIgnoreCase))
			{
				explicitControlId = IDOK;
				label = label.Substring("idok:".Length);
			}
			else if (label.StartsWith("idcancel:", StringComparison.OrdinalIgnoreCase))
			{
				explicitControlId = IDCANCEL;
				label = label.Substring("idcancel:".Length);
			}
			int controlId = explicitControlId != 0 ? explicitControlId : (IsNativeOkButtonLabel(label) ? IDOK : (IsNativeCancelButtonLabel(label) ? IDCANCEL : nextControlId++));
			buttons.Add(new NativeWindowButton(IntPtr.Zero, label, enabled, controlId));
		}
		NativeFamilyScanDialogAction action = ResolveNativeFamilyScanDialogAction(title, body, buttons, activeFamilyEditScope);
		string reason = action == NativeFamilyScanDialogAction.None ? string.Empty : ResolveNativeFamilyScanDialogReason(action, title, body);
		return action.ToString() + "|" + reason;
	}

	private static bool TryClickNativeDialogAction(NativeFamilyScanDialog dialog, ref string clickedButtonText)
	{
		clickedButtonText = string.Empty;
		if (dialog == null || dialog.Handle == IntPtr.Zero)
		{
			return false;
		}
		List<NativeWindowButton> buttons = ((dialog.Buttons != null && dialog.Buttons.Count > 0) ? dialog.Buttons : GetChildButtons(dialog.Handle));
		if (buttons.Count == 0)
		{
			return false;
		}
		NativeWindowButton okButton = buttons.FirstOrDefault([SpecialName] (NativeWindowButton x) => x.Enabled && IsNativeOkButton(x));
		NativeWindowButton cancelButton = buttons.FirstOrDefault([SpecialName] (NativeWindowButton x) => x.Enabled && IsNativeCancelButton(x));
		NativeWindowButton button = null;
		if (dialog.Action == NativeFamilyScanDialogAction.Confirm)
		{
			if (okButton != null && okButton.Handle != IntPtr.Zero)
			{
				button = okButton;
			}
		}
		else if (dialog.Action == NativeFamilyScanDialogAction.Cancel && cancelButton != null && cancelButton.Handle != IntPtr.Zero)
		{
			button = cancelButton;
		}
		if (button == null || button.Handle == IntPtr.Zero)
		{
			if (dialog.Action == NativeFamilyScanDialogAction.Cancel)
			{
				clickedButtonText = "IDCANCEL";
				SendMessage(dialog.Handle, 273, new IntPtr(2), IntPtr.Zero);
				SendMessage(dialog.Handle, 16, IntPtr.Zero, IntPtr.Zero);
				return true;
			}
			return false;
		}
		clickedButtonText = (string.IsNullOrWhiteSpace(button.Text) ? button.Handle.ToString() : button.Text);
		SendMessage(button.Handle, 245, IntPtr.Zero, IntPtr.Zero);
		return true;
	}

	private static bool IsNativeCancelButtonLabel(string text)
	{
		string normalized = NormalizeDialogText(text).Replace("&", string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(normalized))
		{
			return false;
		}
		return ContainsAny(normalized, "cancel", "close", "취소", "닫기");
	}

	private static bool IsNativeCancelButton(NativeWindowButton button)
	{
		if (button == null)
		{
			return false;
		}
		return button.ControlId == 2 || IsNativeCancelButtonLabel(button.Text) || ContainsAny(NormalizeDialogText(button.Text), "취소", "닫기");
	}

	private static bool IsNativeOkButton(NativeWindowButton button)
	{
		if (button == null)
		{
			return false;
		}
		if (IsNativeDeleteButtonLabel(button.Text))
		{
			return false;
		}
		if (IsNativeOkButtonLabel(button.Text) || string.Equals(NormalizeDialogText(button.Text).Trim(), "확인", StringComparison.OrdinalIgnoreCase) || ContainsAny(NormalizeDialogText(button.Text), "계속"))
		{
			return true;
		}
		return button.ControlId == IDOK && string.IsNullOrWhiteSpace(button.Text);
	}

	private static bool IsNativeOkButtonLabel(string text)
	{
		string normalized = NormalizeDialogText(text).Replace("&", string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(normalized))
		{
			return false;
		}
		if (ContainsAny(normalized, "remove", "delete", "제거", "삭제"))
		{
			return false;
		}
		return string.Equals(normalized, "ok", StringComparison.OrdinalIgnoreCase) || string.Equals(normalized, "확인", StringComparison.OrdinalIgnoreCase) || ContainsAny(normalized, "confirm", "continue", "proceed", "계속");
	}

	private static bool IsOpeningNotCuttingAnythingText(string text)
	{
		return ContainsAll(NormalizeDialogText(text), "opening", "not", "cutting", "anything");
	}

	private static bool IsNativeDeleteButtonLabel(string text)
	{
		string normalized = NormalizeDialogText(text).Replace("&", string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(normalized))
		{
			return false;
		}
		return ContainsAny(normalized, "delete instance", "delete type", "delete instances", "delete types", "delete constraints", "remove constraints", "delete", "remove", "삭제", "제거");
	}

	private static bool HasOnlyDeleteOrCancelButtons(IList<NativeWindowButton> buttons)
	{
		List<NativeWindowButton> enabledButtons = (buttons ?? new List<NativeWindowButton>()).Where([SpecialName] (NativeWindowButton x) => x != null && x.Enabled).ToList();
		if (enabledButtons.Count == 0)
		{
			return false;
		}
		bool hasCancel = enabledButtons.Any([SpecialName] (NativeWindowButton x) => IsNativeCancelButton(x));
		bool hasDelete = enabledButtons.Any([SpecialName] (NativeWindowButton x) => IsNativeDeleteButtonLabel(x.Text));
		if (!hasCancel || !hasDelete)
		{
			return false;
		}
		return enabledButtons.All([SpecialName] (NativeWindowButton x) => IsNativeCancelButton(x) || IsNativeDeleteButtonLabel(x.Text));
	}

	private static bool LooksLikeStandardFamilyScanDialog(string text)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return false;
		}
		if (IsDestructiveDeleteChoiceText(text) || HasRemoveOrDeleteConstraintChoice(text))
		{
			return false;
		}
		bool looksLikeWarning = ContainsAny(text, "warning", "warn", "docwarn", "경고", "caution", "error", "failed", "cannot", "can not", "오류", "실패", "할 수 없습니다");
		bool looksLikeFamilyScanIssue = ContainsAny(text, "family", "families", "geometry", "constraint", "constraints", "parameter", "formula", "dimension", "thumbnail", "preview", "image", "패밀리", "형상", "기하", "구속", "제약", "매개변수", "치수", "미리보기", "이미지");
		return looksLikeWarning && looksLikeFamilyScanIssue;
	}

	private static string BuildNativeButtonSummary(IEnumerable<NativeWindowButton> buttons)
	{
		List<string> parts = new List<string>();
		foreach (NativeWindowButton button in (buttons ?? Enumerable.Empty<NativeWindowButton>()).Where([SpecialName] (NativeWindowButton x) => x != null))
		{
			parts.Add((button.Enabled ? "enabled" : "disabled") + ":" + button.ControlId.ToString(CultureInfo.InvariantCulture) + ":" + (string.IsNullOrWhiteSpace(button.Text) ? "(blank)" : button.Text.Trim()));
		}
		return string.Join(" | ", parts);
	}

	private static List<string> GetDescendantTexts(IntPtr parentHandle)
	{
		List<string> list = new List<string>();
		EnumChildWindows(parentHandle, [SpecialName] (IntPtr childHandle, IntPtr lParam) =>
		{
			string windowTextSafe = GetWindowTextSafe(childHandle);
			if (!string.IsNullOrWhiteSpace(windowTextSafe))
			{
				list.Add(windowTextSafe.Trim());
			}
			return true;
		}, IntPtr.Zero);
		return list;
	}

	private static List<NativeWindowButton> GetChildButtons(IntPtr parentHandle)
	{
		List<NativeWindowButton> list = new List<NativeWindowButton>();
		EnumChildWindows(parentHandle, [SpecialName] (IntPtr childHandle, IntPtr lParam) =>
		{
			if (!IsWindowVisible(childHandle))
			{
				return true;
			}
			if (GetWindowClassName(childHandle).IndexOf("Button", StringComparison.OrdinalIgnoreCase) < 0)
			{
				return true;
			}
			list.Add(new NativeWindowButton(childHandle, GetWindowTextSafe(childHandle), IsWindowEnabled(childHandle), GetDlgCtrlID(childHandle)));
			return true;
		}, IntPtr.Zero);
		return list;
	}

	private static string GetWindowClassName(IntPtr hWnd)
	{
		StringBuilder buffer = new StringBuilder(256);
		GetClassName(hWnd, buffer, buffer.Capacity);
		return buffer.ToString();
	}

	private static string GetWindowTextSafe(IntPtr hWnd)
	{
		int length = GetWindowTextLength(hWnd);
		if (length <= 0)
		{
			return string.Empty;
		}
		StringBuilder buffer = new StringBuilder(checked(length + 1));
		GetWindowText(hWnd, buffer, buffer.Capacity);
		return buffer.ToString();
	}

	[DllImport("user32.dll")]
	private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

	[DllImport("user32.dll")]
	private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

	[DllImport("user32.dll")]
	private static extern bool IsWindowVisible(IntPtr hWnd);

	[DllImport("user32.dll")]
	private static extern bool IsWindowEnabled(IntPtr hWnd);

	[DllImport("user32.dll")]
	private static extern int GetDlgCtrlID(IntPtr hWnd);

	[DllImport("user32.dll")]
	private static extern uint GetWindowThreadProcessId(IntPtr hWnd, ref uint processId);

	[DllImport("user32.dll", CharSet = CharSet.Auto)]
	private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

	[DllImport("user32.dll", CharSet = CharSet.Auto)]
	private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

	[DllImport("user32.dll", CharSet = CharSet.Auto)]
	private static extern int GetWindowTextLength(IntPtr hWnd);

	[DllImport("user32.dll")]
	private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

	public void Dispose()
	{
		if (_disposed)
		{
			return;
		}
		_disposed = true;
		if (_nativeDialogTimer != null)
		{
			try
			{
				_nativeDialogTimer.Dispose();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		if (_uiApplication != null)
		{
			try
			{
				_uiApplication.DialogBoxShowing -= HandleDialogBoxShowing;
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
		}
	}

	void IDisposable.Dispose()
	{
		//ILSpy generated this explicit interface implementation from .override directive in Dispose
		this.Dispose();
	}
}
