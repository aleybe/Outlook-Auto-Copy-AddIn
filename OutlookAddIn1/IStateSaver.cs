namespace OutlookAddIn1
{
    public interface IStateSaver
    {
        bool IsEnabled { get; set; }
        void Save();
        void Load();
    }
}