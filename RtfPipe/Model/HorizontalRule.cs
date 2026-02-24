namespace RtfPipe.Model
{
  internal class HorizontalRule : Node
  {
    internal override void Visit(INodeVisitor visitor)
    {
      visitor.Visit(this);
    }
  }
}
