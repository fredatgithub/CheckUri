using Microsoft.VisualStudio.TestTools.UnitTesting;
using PreClic.Helpers;

namespace UnitTestProjectPreClic
{
  [TestClass]
  public class UnitTestPreClic
  {
    [TestMethod]
    public void TestMethod_TrimStart_http_true()
    {
      const string source = "http://myserver.com/rep1";
      const string source2 = "http://";
      const string expected = "myserver.com/rep1";
      string result = StringHelpers.TrimStart(source, source2);
      Assert.AreEqual(result, expected);
    }

    [TestMethod]
    public void TestMethod_TrimStart_http_false()
    {
      const string source = "https://myserver.com/rep1";
      const string source2 = "http://";
      const string expected = "https://myserver.com/rep1";
      string result = StringHelpers.TrimStart(source, source2);
      Assert.AreEqual(result, expected);
    }

    [TestMethod]
    public void TestMethod_TrimStart_https_true()
    {
      const string source = "https://myserver.com/rep1";
      const string source2 = "https://";
      const string expected = "myserver.com/rep1";
      string result = StringHelpers.TrimStart(source, source2);
      Assert.AreEqual(result, expected);
    }

    [TestMethod]
    public void TestMethod_TrimStart_https_false()
    {
      const string source = "http://myserver.com/rep1";
      const string source2 = "https://";
      const string expected = "http://myserver.com/rep1";
      string result = StringHelpers.TrimStart(source, source2);
      Assert.AreEqual(result, expected);
    }

    [TestMethod]
    public void TestMethod_GetServerNameFromUrl()
    {
      const string source = "http://myserver.com/rep1";
      const string expected = "myserver.com";
      string result = StringHelpers.GetServerNameFromUrl(source);
      Assert.AreEqual(result, expected);
    }

    [TestMethod]
    public void TestMethod_GetServerNameFromUrlAndPort()
    {
      const string source = "http://myserver.com:8080/rep1";
      const string expected = "myserver.com";
      string result = StringHelpers.GetServerNameFromUrl(source);
      Assert.AreEqual(result, expected);
    }
  }
}