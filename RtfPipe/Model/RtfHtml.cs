using System.Collections.Generic;

namespace RtfPipe.Model
{
  internal class RtfHtml
  {
    public UnitValue DefaultTabWidth { get; set; }
    public Dictionary<IToken, object> Metadata { get; set; }
    public Element Root { get; set; }
  }
}
