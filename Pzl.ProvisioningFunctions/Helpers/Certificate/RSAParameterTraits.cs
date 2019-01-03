using System;

namespace Pzl.ProvisioningFunctions.Helpers.Certificate
{
    internal class RSAParameterTraits
    {
        public int SizeD = -1;
        public int SizeDp = -1;
        public int SizeDq = -1;
        public int SizeExp = -1;
        public int SizeInvQ = -1;

        public int SizeMod = -1;
        public int SizeP = -1;
        public int SizeQ = -1;

        public RSAParameterTraits(int modulusLengthInBits)
        {
            // The modulus length is supposed to be one of the common lengths, which is the commonly referred to strength of the key,
            // like 1024 bit, 2048 bit, etc.  It might be a few bits off though, since if the modulus has leading zeros it could show
            // up as 1016 bits or something like that.
            int assumedLength = -1;
            double logbase = Math.Log(modulusLengthInBits, 2);
            if (logbase == (int) logbase)
            {
                // It's already an even power of 2
                assumedLength = modulusLengthInBits;
            }
            else
            {
                // It's not an even power of 2, so round it up to the nearest power of 2.
                assumedLength = (int) (logbase + 1.0);
                assumedLength = (int) Math.Pow(2, assumedLength);
                // you should verify that this really does the 'right' thing!
            }

            switch (assumedLength)
            {
                case 1024:
                    SizeMod = 0x80;
                    SizeExp = -1;
                    SizeD = 0x80;
                    SizeP = 0x40;
                    SizeQ = 0x40;
                    SizeDp = 0x40;
                    SizeDq = 0x40;
                    SizeInvQ = 0x40;
                    break;
                case 2048:
                    SizeMod = 0x100;
                    SizeExp = -1;
                    SizeD = 0x100;
                    SizeP = 0x80;
                    SizeQ = 0x80;
                    SizeDp = 0x80;
                    SizeDq = 0x80;
                    SizeInvQ = 0x80;
                    break;
                case 4096:
                    SizeMod = 0x200;
                    SizeExp = -1;
                    SizeD = 0x200;
                    SizeP = 0x100;
                    SizeQ = 0x100;
                    SizeDp = 0x100;
                    SizeDq = 0x100;
                    SizeInvQ = 0x100;
                    break;
                default:
                    break;
            }
        }
    }
}