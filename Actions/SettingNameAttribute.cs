namespace Actions
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct)]
    public class SettingNameAttribute(string name) : Attribute
    {
        public string Name { get; set;} = name;
    }
}