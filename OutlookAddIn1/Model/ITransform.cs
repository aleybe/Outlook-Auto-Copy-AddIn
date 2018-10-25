namespace ConsoleApp4.Model
{
    // This is a tiny helper interface I use when dealing with converting or transforming one piece of data into another.
    public interface IConvert<TFrom, TResult>
    {
        TResult ConvertTo(TFrom data);
        TFrom ConvertFrom(TResult data);
    }
}