namespace ProjectReport.Actions
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct)]
    public class SettingNameAttribute : Attribute
    {
        public string Name { get; }

        public SettingNameAttribute(string name)
        {
            Name = name;
        }
    }
}