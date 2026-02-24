namespace RtfPipe
{
  public class BinaryToken : IToken
  {
    public byte[] Value { get; set; }
    public TokenType Type => TokenType.Text;
  }
}
