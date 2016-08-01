using System;
using System.IO;
using PreClic.Properties;

namespace PreClic.Helpers
{
  public static class StringHelpers
  {
    public static string TrimStart(string originalString, string stringToBeTrimmed)
    {
      return originalString.Substring(0, stringToBeTrimmed.Length) == stringToBeTrimmed ? originalString.Substring(stringToBeTrimmed.Length) : originalString;
    }

    public static string GetServerNameFromUrl(string url)
    {
      string result = string.Empty;
      result = TrimStart(url, "http://");
      result = TrimStart(result, "https://");
      result = result.Substring(0, result.IndexOf("/", StringComparison.InvariantCulture));
      // remove port number if it exists
      result = result.Substring(0, result.IndexOf(":", StringComparison.InvariantCulture));
      return result;
    }

    public static bool IniFileExists(string fileName)
    {
      bool result = false;
      string DirectoryFileName = Path.GetDirectoryName(fileName);
      if (DirectoryFileName != null)
      {
        result = File.Exists(Path.Combine(DirectoryFileName, Settings.Default.SchemaIniFileName));
      }

      return result;
    }

    public static string ExtractFileName(string longFileName)
    {
      return Path.GetFileName(longFileName);
    }

    public static string GetDirectoryName(string longFileName)
    {
      return Path.GetDirectoryName(longFileName);
    }

    public static string AddBackSlash(string fileName)
    {
      return fileName.EndsWith("\\") ? fileName : fileName + "\\";
    }

    public static string AddSlash(string fileName)
    {
      return fileName.EndsWith("/") ? fileName : fileName + "/";
    }
  }
}