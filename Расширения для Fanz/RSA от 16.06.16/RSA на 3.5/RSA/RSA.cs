using System;
using System.Collections.Generic;

namespace RSA
{
    public class RSAlib
    {
        private byte p;        
        private byte q;         
        private long phi;     
        private long n;
        private long e;
        private long d;

        private struct ExtendedEuclideanResult
        {
            public long u1;
            public long u2;
            public long gcd;
        }

        public RSAlib()
        {
 
        }

        private void InitKeyData()
        {
            Random random = new Random();

            byte[] simple = GetNotDivideable();
            p = simple[random.Next(10, simple.Length)];
            q = simple[random.Next(10, simple.Length)];
            n = (p * q);
            phi = (p - 1) * (q - 1);
            List<long> possibleE = GetAllPossibleE(phi);

            do
            {
                e = possibleE[random.Next(10, possibleE.Count)];
                d = ExtendedEuclide(e % phi, phi).u1;
            } while (d < 0);
        }

        public long GetNKey()
        {
            return n;
        }

        public long GetDKey()
        {
            return d;
        }

        public string encode(string t)
        {
            InitKeyData();
            string text = t;
            string outStr = "";
            System.Text.UTF8Encoding enc = new System.Text.UTF8Encoding();
            byte[] strBytes = enc.GetBytes(text);
            foreach (byte value in strBytes)
            {
                long encryptedValue = ModuloPow(value, e, n);
                outStr += encryptedValue + "|";
            }

            return outStr;
        }

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
        static List<long> GetAllPossibleE(long phi)
        {
            List<long> result = new List<long>();

            for (long i = 2; i < phi; i++)
            {
                if (ExtendedEuclide(i, phi).gcd == 1)
                {
                    result.Add(i);
                }
            }

            return result;
        }
        private static ExtendedEuclideanResult ExtendedEuclide(long a, long b)
        {
            long u1 = 1;
            long u3 = a;
            long v1 = 0;
            long v3 = b;

            while (v3 > 0)
            {
                long q0 = u3 / v3;
                long q1 = u3 % v3;

                long tmp = v1 * q0;
                long tn = u1 - tmp;
                u1 = v1;
                v1 = tn;

                u3 = v3;
                v3 = q1;
            }

            long tmp2 = u1 * (a);
            tmp2 = u3 - (tmp2);
            long res = tmp2 / (b);

            ExtendedEuclideanResult result = new ExtendedEuclideanResult()
            {
                u1 = u1,
                u2 = res,
                gcd = u3
            };

            return result;
        }

        static private byte[] GetNotDivideable()
        {
            List<byte> notDivideable = new List<byte>();

            for (int x = 2; x < 256; x++)
            {
                int n = 0;
                for (int y = 1; y <= x; y++)
                {
                    if (x % y == 0)
                        n++;
                }

                if (n <= 2)
                    notDivideable.Add((byte)x);
            }
            return notDivideable.ToArray();
        }

    }
}
