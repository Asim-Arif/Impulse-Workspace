using System.Text;

namespace Impulse
{
    public class Encryption
    {

        private byte[] S = new byte[256];
        private byte[] K = new byte[256];
        public string CryptRC4(string text, string key)
        {
            Initialize(key);
            byte[] data = Encoding.GetEncoding(1252).GetBytes(text); //Encoding.Default.GetBytes(text); //Encoding.ASCII.GetBytes(text);
            byte[] cipher = new byte[data.Length];

            int i = 0;
            int j = 0;

            for (int l = 0; l < data.Length; l++)
            {
                i = (i + 1) % 256;
                j = (j + S[i]) % 256;

                Swap(S, i, j);
                int t = (S[i] + S[j]) % 256;
                cipher[l] = (byte)(data[l] ^ S[t]);
            }
            return Encoding.Default.GetString(cipher); //Encoding.ASCII.GetString(cipher);
        }
        private void Initialize(string key)
        {
            for (int i = 0; i < 256; i++)
            {
                S[i] = (byte)i; 
                K[i] = (byte)key[i % key.Length];
            }

            int j = 0;
            for (int i = 0; i < 256; i++)
            {
                j = (j + S[i] + K[i]) % 256;
                Swap(S, i, j);
            }
        }
        private void Swap(byte[] array, int i, int j)
        {
            byte temp = array[i];
            array[i] = array[j];
            array[j] = temp;
        }
        public string ToHexDump(string text)
        {
            StringBuilder hex = new StringBuilder(text.Length * 2);
            foreach (char c in text)
            {
                hex.AppendFormat("{0:x2}", (int)c);
            }
            return hex.ToString();
        }
        public string FromHexDump(string hex)
        {
            //StringBuilder text = new StringBuilder(hex.Length / 2);
            //for (int i = 0; i < hex.Length; i += 2)
            //{
            //    string h = hex.Substring(i, 2);
            //    if (int.TryParse(h, System.Globalization.NumberStyles.HexNumber, null, out int value))
            //    {
            //        text.Append((char)value);
            //    }
            //    else 
            //    {
            //        throw new FormatException("Invalid hexadecimal string format.");
            //    }

            //}
            //return text.ToString();
            return hex;
        }
    }
}
