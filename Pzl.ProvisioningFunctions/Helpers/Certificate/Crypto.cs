using System;
using System.IO;
using System.Security.Cryptography;

namespace Cumulus.Monads.Helpers.Certificate
{
    public class Crypto
    {
        /// <summary>
        ///     This helper function parses an RSA private key using the ASN.1 format
        /// </summary>
        /// <param name="privateKeyBytes">Byte array containing PEM string of private key.</param>
        /// <returns>
        ///     An instance of <see cref="RSACryptoServiceProvider" /> rapresenting the requested private key.
        ///     Null if method fails on retriving the key.
        /// </returns>
        public static RSACryptoServiceProvider DecodeRsaPrivateKey(byte[] privateKeyBytes)
        {
            var ms = new MemoryStream(privateKeyBytes);
            var rd = new BinaryReader(ms);

            try
            {
                ushort shortValue = rd.ReadUInt16();

                switch (shortValue)
                {
                    case 0x8130:
                        // If true, data is little endian since the proper logical seq is 0x30 0x81
                        rd.ReadByte(); //advance 1 byte
                        break;
                    case 0x8230:
                        rd.ReadInt16(); //advance 2 bytes
                        break;
                    default:
                        return null;
                }

                shortValue = rd.ReadUInt16();
                if (shortValue != 0x0102) // (version number)
                    return null;

                byte byteValue = rd.ReadByte();
                if (byteValue != 0x00)
                    return null;

                // The data following the version will be the ASN.1 data itself, which in our case
                // are a sequence of integers.

                var parms = new CspParameters
                {
                    Flags = CspProviderFlags.UseMachineKeyStore,
                    KeyContainerName = Guid.NewGuid().ToString().ToUpperInvariant()
                };

                var rsa = new RSACryptoServiceProvider(parms);
                var rsAparams = new RSAParameters {Modulus = rd.ReadBytes(Helpers.DecodeIntegerSize(rd))};


                // Argh, this is a pain.  From emperical testing it appears to be that RSAParameters doesn't like byte buffers that
                // have their leading zeros removed.  The RFC doesn't address this area that I can see, so it's hard to say that this
                // is a bug, but it sure would be helpful if it allowed that. So, there's some extra code here that knows what the
                // sizes of the various components are supposed to be.  Using these sizes we can ensure the buffer sizes are exactly
                // what the RSAParameters expect.  Thanks, Microsoft.
                var traits = new RSAParameterTraits(rsAparams.Modulus.Length * 8);

                rsAparams.Modulus = Helpers.AlignBytes(rsAparams.Modulus, traits.SizeMod);
                rsAparams.Exponent = Helpers.AlignBytes(rd.ReadBytes(Helpers.DecodeIntegerSize(rd)), traits.SizeExp);
                rsAparams.D = Helpers.AlignBytes(rd.ReadBytes(Helpers.DecodeIntegerSize(rd)), traits.SizeD);
                rsAparams.P = Helpers.AlignBytes(rd.ReadBytes(Helpers.DecodeIntegerSize(rd)), traits.SizeP);
                rsAparams.Q = Helpers.AlignBytes(rd.ReadBytes(Helpers.DecodeIntegerSize(rd)), traits.SizeQ);
                rsAparams.DP = Helpers.AlignBytes(rd.ReadBytes(Helpers.DecodeIntegerSize(rd)), traits.SizeDp);
                rsAparams.DQ = Helpers.AlignBytes(rd.ReadBytes(Helpers.DecodeIntegerSize(rd)), traits.SizeDq);
                rsAparams.InverseQ = Helpers.AlignBytes(rd.ReadBytes(Helpers.DecodeIntegerSize(rd)), traits.SizeInvQ);

                rsa.ImportParameters(rsAparams);
                return rsa;
            }
            //catch (Exception e)
            //{
            //    return null;
            //}
            finally
            {
                rd.Close();
            }
        }
    }
}