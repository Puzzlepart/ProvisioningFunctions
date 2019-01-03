using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace Pzl.ProvisioningFunctions.Helpers.Certificate
{
    public class Certificate
    {
        public Certificate(string cert, string key, string password)
        {
            PublicCertificate = cert.Replace("\r\n", string.Empty).Replace("\n", string.Empty).Replace("\r", string.Empty);
            PrivateKey = key.Replace("\r\n", string.Empty).Replace("\n", string.Empty).Replace("\r", string.Empty);
            Password = password;
        }

        public X509Certificate2 GetCertificateFromPEMstring(bool certOnly)
        {
            if (certOnly)
                return GetCertificateFromPEMstring(PublicCertificate);
            return GetCertificateFromPEMstring(PublicCertificate, PrivateKey, Password);
        }

        public static X509Certificate2 GetCertificateFromPEMstring(string publicCert)
        {
            return new X509Certificate2(Encoding.UTF8.GetBytes(publicCert));
        }

        public static X509Certificate2 GetCertificateFromPEMstring(string publicCert, string privateKey,
            string password)
        {
            var certBuffer = Helpers.GetBytesFromPEM(publicCert, PemStringType.Certificate);
            var keyBuffer = Helpers.GetBytesFromPEM(privateKey, PemStringType.RsaPrivateKey);

            var certificate = new X509Certificate2(certBuffer, password);

            var prov = Crypto.DecodeRsaPrivateKey(keyBuffer);
            certificate.PrivateKey = prov;

            return certificate;
        }

        #region Fields

        #endregion

        #region Properties

        public string PublicCertificate { get; set; }

        public string PrivateKey { get; set; }

        public string Password { get; set; }

        #endregion
    }
}