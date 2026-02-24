namespace RtfPipe
{
  public class ListStyleReference
  {
    public int Id { get; }
    public ListStyle Style { get; }

    internal ListStyleReference(int id, ListStyle style)
    {
      Id = id;
      Style = style;
    }
  }
}
