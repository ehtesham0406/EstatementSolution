using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Security.Cryptography;

namespace QCash.EStatement.BAL
{
    
   public static class XmlEncryptor
    {
        private static readonly string key = "Z7k9P2sX8rV1bQ3nH5mL0tF6dR4wC1yA"; // 32 chars = AES-256

        public static void EncryptXml(string inputFile, string outputFile)
        {
            byte[] plainBytes = File.ReadAllBytes(inputFile);
            byte[] keyBytes = Encoding.UTF8.GetBytes(key);
            byte[] iv = new byte[16]; // AES IV

            using (Aes aes = Aes.Create())
            {
                aes.Key = keyBytes;
                aes.IV = iv;
                using (var encryptor = aes.CreateEncryptor())
                using (var fs = new FileStream(outputFile, FileMode.Create))
                using (var cs = new CryptoStream(fs, encryptor, CryptoStreamMode.Write))
                {
                    cs.Write(plainBytes, 0, plainBytes.Length);
                }
            }
        }

        public static string DecryptXml(string filePath)
        {
            byte[] encryptedBytes = File.ReadAllBytes(filePath);
            byte[] keyBytes = Encoding.UTF8.GetBytes(key);
            byte[] iv = new byte[16];

            using (Aes aes = Aes.Create())
            {
                aes.Key = keyBytes;
                aes.IV = iv;
                using (var decryptor = aes.CreateDecryptor())
                using (var ms = new MemoryStream(encryptedBytes))
                using (var cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                using (var sr = new StreamReader(cs))
                {
                    return sr.ReadToEnd();
                }
            }
        }
    }
}
