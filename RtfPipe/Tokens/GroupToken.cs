namespace RtfPipe
{
  public class GroupEnd : IToken
  {
    public TokenType Type => TokenType.GroupEnd;

    public override string ToString()
    {
      return "}";
    }
  }
}
