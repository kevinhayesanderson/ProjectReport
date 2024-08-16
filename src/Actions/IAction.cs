namespace Actions
{
    public interface IAction
    {
        void Init();

        bool Run();

        bool Validate();
    }
}