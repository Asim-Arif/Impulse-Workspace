using System;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Impulse.Models.License;

namespace Impulse.Services.Licensing
{
    public static class LicenseCrypto
    {
        private const string LicensePrefix = "IMPULSE-LIC-";

        public static (bool IsValid, LicensePayload? Payload, string ErrorMessage) VerifyLicense(string licenseString, string publicKeyXml)
        {
            if (string.IsNullOrWhiteSpace(licenseString))
            {
                return (false, null, "License key is empty.");
            }

            try
            {
                var trimmed = licenseString.Trim();
                if (trimmed.StartsWith(LicensePrefix, StringComparison.OrdinalIgnoreCase))
                {
                    trimmed = trimmed.Substring(LicensePrefix.Length);
                }

                var parts = trimmed.Split('.');
                if (parts.Length != 2)
                {
                    return (false, null, "Malformed license string format.");
                }

                var payloadBytes = Convert.FromBase64String(parts[0]);
                var signatureBytes = Convert.FromBase64String(parts[1]);

                using var rsa = RSA.Create();
                rsa.FromXmlString(publicKeyXml);

                var isSignatureValid = rsa.VerifyData(
                    payloadBytes,
                    signatureBytes,
                    HashAlgorithmName.SHA256,
                    RSASignaturePadding.Pkcs1);

                if (!isSignatureValid)
                {
                    return (false, null, "Digital signature verification failed. License has been tampered with or corrupted.");
                }

                var json = Encoding.UTF8.GetString(payloadBytes);
                var payload = JsonSerializer.Deserialize<LicensePayload>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                if (payload == null)
                {
                    return (false, null, "Unable to deserialize license payload.");
                }

                return (true, payload, string.Empty);
            }
            catch (Exception ex)
            {
                return (false, null, $"License verification error: {ex.Message}");
            }
        }

        public static string GenerateLicense(LicensePayload payload, string privateKeyXml)
        {
            var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = false });
            var payloadBytes = Encoding.UTF8.GetBytes(json);

            using var rsa = RSA.Create();
            rsa.FromXmlString(privateKeyXml);

            var signatureBytes = rsa.SignData(
                payloadBytes,
                HashAlgorithmName.SHA256,
                RSASignaturePadding.Pkcs1);

            var payloadBase64 = Convert.ToBase64String(payloadBytes);
            var signatureBase64 = Convert.ToBase64String(signatureBytes);

            return $"{LicensePrefix}{payloadBase64}.{signatureBase64}";
        }
    }
}
