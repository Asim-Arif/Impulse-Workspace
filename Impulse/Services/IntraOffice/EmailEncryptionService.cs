using System;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Extensions.Configuration;

namespace Impulse.Services.IntraOffice
{
    public interface IEmailEncryptionService
    {
        string Encrypt(string plainText);
        string Decrypt(string cipherText);
    }

    public class EmailEncryptionService : IEmailEncryptionService
    {
        private readonly byte[] _key;

        public EmailEncryptionService(IConfiguration configuration)
        {
            var secret = configuration["EmailEncryption:SecretKey"] ?? "IntraComCRM-Enterprise-SmtpSecretKey-2026-SecurityKey!";
            using var sha = SHA256.Create();
            _key = sha.ComputeHash(Encoding.UTF8.GetBytes(secret));
        }

        public string Encrypt(string plainText)
        {
            if (string.IsNullOrEmpty(plainText))
                return string.Empty;

            using var aes = Aes.Create();
            aes.Key = _key;
            aes.GenerateIV();

            using var encryptor = aes.CreateEncryptor(aes.Key, aes.IV);
            var plainBytes = Encoding.UTF8.GetBytes(plainText);
            var cipherBytes = encryptor.TransformFinalBlock(plainBytes, 0, plainBytes.Length);

            var combinedBytes = new byte[aes.IV.Length + cipherBytes.Length];
            Buffer.BlockCopy(aes.IV, 0, combinedBytes, 0, aes.IV.Length);
            Buffer.BlockCopy(cipherBytes, 0, combinedBytes, aes.IV.Length, cipherBytes.Length);

            return Convert.ToBase64String(combinedBytes);
        }

        public string Decrypt(string cipherText)
        {
            if (string.IsNullOrEmpty(cipherText))
                return string.Empty;

            try
            {
                var combinedBytes = Convert.FromBase64String(cipherText);
                if (combinedBytes.Length < 16)
                    return string.Empty;

                using var aes = Aes.Create();
                aes.Key = _key;

                var iv = new byte[16];
                Buffer.BlockCopy(combinedBytes, 0, iv, 0, 16);
                aes.IV = iv;

                var cipherBytes = new byte[combinedBytes.Length - 16];
                Buffer.BlockCopy(combinedBytes, 16, cipherBytes, 0, cipherBytes.Length);

                using var decryptor = aes.CreateDecryptor(aes.Key, aes.IV);
                var decryptedBytes = decryptor.TransformFinalBlock(cipherBytes, 0, cipherBytes.Length);

                return Encoding.UTF8.GetString(decryptedBytes);
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
