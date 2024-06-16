namespace Actions
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct)]
    public class ActionNameAttribute(string name) : Attribute
    {
        public string Name => name;
    }
}