using System.Collections.Generic;

public sealed class FamilyBrowserStandardLibrarySlot
{
	public string SlotKey { get; set; }

	public string Discipline { get; set; }

	public string DisplayName { get; set; }

	public string RegistrationPath { get; set; }

	public string SourceId { get; set; }

	public string StandardRvtPath { get; set; }

	public string SnapshotPath { get; set; }

	public string StandardListPath { get; set; }

	public string StandardListSheetName { get; set; }

	public string LastSnapshotAtUtc { get; set; }

	public bool Enabled { get; set; }

	public FamilyBrowserStandardLibrarySlot()
	{
		SlotKey = string.Empty;
		Discipline = string.Empty;
		DisplayName = string.Empty;
		RegistrationPath = string.Empty;
		SourceId = string.Empty;
		StandardRvtPath = string.Empty;
		SnapshotPath = string.Empty;
		StandardListPath = string.Empty;
		StandardListSheetName = string.Empty;
		LastSnapshotAtUtc = string.Empty;
		Enabled = true;
	}

	public static FamilyBrowserStandardLibrarySlot CreateIntegrated()
	{
		return new FamilyBrowserStandardLibrarySlot
		{
			SlotKey = "integrated",
			Discipline = "Integrated",
			DisplayName = "Integrated Standard RVT",
			Enabled = true
		};
	}

	public static List<FamilyBrowserStandardLibrarySlot> CreateDefaultDisciplines()
	{
		return new List<FamilyBrowserStandardLibrarySlot>
		{
			CreateDiscipline("Architecture", "Architecture"),
			CreateDiscipline("Structure", "Structure"),
			CreateDiscipline("Mechanical", "Mechanical"),
			CreateDiscipline("Electrical", "Electrical"),
			CreateDiscipline("FireProtection", "Fire Protection"),
			CreateDiscipline("Other", "Other")
		};
	}

	public static FamilyBrowserStandardLibrarySlot CreateDiscipline(string discipline, string displayName)
	{
		return new FamilyBrowserStandardLibrarySlot
		{
			SlotKey = "discipline-" + FamilyBrowserPolicyKey.Normalize(discipline),
			Discipline = discipline,
			DisplayName = displayName,
			Enabled = true
		};
	}
}
