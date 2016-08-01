namespace PreClic
{
  internal class Line
  {
    public string EnvironnmentName { get; set; }
    public string UrlString { get; set; }
    public bool PingOk { get; set; }
    public bool HttpReachable { get; set; }
    public bool CanLogin { get; set; }
  }
}