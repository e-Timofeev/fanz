namespace RSAdelib
{
    public class Delib
    {
        public string decode(string text, string n_s, string d_s)
        {
            string outStr = "";
            long n = long.Parse(n_s);
            long d = long.Parse(d_s);
            long[] arr = GetDecArrayFromText(text);
            byte[] bytes = new byte[arr.Length];
            System.Text.UTF8Encoding enc = new System.Text.UTF8Encoding();
            int j = 0;
            foreach (int i in arr)
            {
                byte decryptedValue = (byte)ModuloPow(i, d, n);

                bytes[j] = decryptedValue;
                j++;

            }
            outStr += enc.GetString(bytes);
            return outStr;
        }

        private long[] GetDecArrayFromText(string text)
        {
            int i = 0;
            foreach (char c in text)
            {
                if (c == '|')
                {
                    i++;
                }
            }

            long[] result = new long[i];
            i = 0;

            string tmp = "";

            foreach (char c in text)
            {
                if (c != '|')
                {
                    tmp += c;
                }
                else
                {
                    result[i] = long.Parse(tmp);
                    i++;
                    tmp = "";
                }
            }

            return result;
        }

        static long ModuloPow(long value, long pow, long modulo)
        {
            long result = value;
            for (int i = 0; i < pow - 1; i++)
            {
                result = (result * value) % modulo;
            }
            return result;
        }

    }
}
