using System;
using System.IO;
using System.Net;

namespace ConsoleApplicationForTest
{
  internal class Program
  {
    private static void Main()
    {
      string[] urls = {
        "http://www.google.fr",
        "http://www.google.com",
        "http://www.google.co.uk",
        "http://www.google.fr",
        "http://www.google.fr"
        };

      foreach (string url in urls)
      {
        //Thread.Sleep(1000);
        string httpResult = GetHttpResponse(url);
        if (httpResult.Contains("ID utilisateur") && httpResult.Contains("Mot de passe"))
        {
          Console.WriteLine("ok");
        }
        else
        {
          Console.WriteLine("KO : " + httpResult);
        }
      }

      Console.WriteLine("Press any key to exit:");
      Console.ReadKey();
    }

    public static string GetHttpResponse(string URL)
    {
      HttpWebRequest request = WebRequest.Create(URL) as HttpWebRequest;
      request.KeepAlive = false;
      request.Timeout = 5000;
      request.Proxy = null;

      request.ServicePoint.ConnectionLeaseTimeout = 5000;
      request.ServicePoint.MaxIdleTime = 5000;

      // Read stream
      string responseString = string.Empty;
      try
      {
        using (WebResponse response = request.GetResponse())
        {
          using (Stream objStream = response.GetResponseStream())
          {
            using (StreamReader objReader = new StreamReader(objStream))
            {
              responseString = objReader.ReadToEnd();
              objReader.Close();
            }
            objStream.Flush();
            objStream.Close();
          }
          response.Close();
        }
      }
      catch (WebException exception)
      {
        // ignored
        //MessageBox.Show("erreur" + exception.Message);
        //Console.WriteLine("erreur : " + exception.Message);
        responseString = "erreur : " + exception.Message;
      }
      finally
      {
        request.Abort();
      }

      return responseString;
    }
  }
}