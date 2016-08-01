using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace PreClic
{
  internal class Functions
  {
    public static byte[] Key = new byte[24]
    {
      1,
      2,
      3,
      4,
      5,
      6,
      7,
      8,
      9,
      10,
      11,
      12,
      13,
      14,
      15,
      16,
      17,
      18,
      19,
      20,
      21,
      22,
      23,
      24
    };
    public static byte[] Iv = new byte[8]
    {
      65,
      110,
      68,
      26,
      69,
      178,
      200,
      219
    };

    public static string Encrypt(string plainText)
    {
      if (plainText.Length <= 0) return string.Empty;
      byte[] bytes = new UTF8Encoding().GetBytes(plainText);
      ICryptoTransform encryptor = new TripleDESCryptoServiceProvider().CreateEncryptor(Key, Iv);
      MemoryStream memoryStream = new MemoryStream();
      CryptoStream cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write);
      cryptoStream.Write(bytes, 0, bytes.Length);
      cryptoStream.FlushFinalBlock();
      memoryStream.Position = 0L;
      byte[] buffer = new byte[checked((int)(memoryStream.Length - 1L) + 1)];
      memoryStream.Read(buffer, 0, checked((int)memoryStream.Length));
      cryptoStream.Close();
      string str = string.Empty;
      int num1 = 0;
      int num2 = checked(buffer.Length - 1);
      int index = num1;
      while (index <= num2)
      {
        str += string.Format("{0:000}", buffer[index]);
        checked { ++index; }
      }

      return str;
    }

    public static string Decrypt(string resultat)
    {
      if (string.Compare(resultat, "", false, CultureInfo.InvariantCulture) == 0) return string.Empty;
      string str;
      try
      {
        byte[] buffer = new byte[checked((int)Math.Round(resultat.Length / 3.0) - 1 + 1)];
        int num1 = 0;
        int num2 = checked(resultat.Length - 1);
        int startIndex = num1;
        while (startIndex <= num2)
        {
          // buffer[checked((int)Math.Round(unchecked((double)startIndex / 3.0)))] = checked((byte)Conversions.ToInteger(resultat.Substring(startIndex, 3)));
          buffer[checked((int)Math.Round(startIndex / 3.0))] = checked((byte)int.Parse(resultat.Substring(startIndex, 3)));
          checked { startIndex += 3; }
        }

        var utF8Encoding = new UTF8Encoding();
        ICryptoTransform decryptor = new TripleDESCryptoServiceProvider().CreateDecryptor(Key, Iv);
        var memoryStream = new MemoryStream();
        var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Write);
        cryptoStream.Write(buffer, 0, buffer.Length);
        cryptoStream.FlushFinalBlock();
        memoryStream.Position = 0L;
        var numArray = new byte[checked((int)(memoryStream.Length - 1L) + 1)];
        memoryStream.Read(numArray, 0, checked((int)memoryStream.Length));
        cryptoStream.Close();
        str = new UTF8Encoding().GetString(numArray);
      }
      catch (Exception exception)
      {
        MessageBox.Show(@"La valeur du mot de passe ne respecte pas la norme triple DES
----------------------------------------------------------------------------
" + exception, @"Error");
        str = string.Empty;
      }

      return str;
    }
  }
}