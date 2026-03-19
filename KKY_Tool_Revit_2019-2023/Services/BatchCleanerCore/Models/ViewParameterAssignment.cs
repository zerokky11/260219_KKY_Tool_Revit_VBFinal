namespace KKY_Tool_Revit.Models
{
    public class ViewParameterAssignment
    {
        public bool Enabled { get; set; }
        public string ParameterName { get; set; }
        public string ParameterValue { get; set; }

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(ParameterName)
                ? string.Empty
                : ParameterName + " = " + ParameterValue;
        }
    }
}
