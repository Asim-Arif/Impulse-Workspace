using System;

namespace Impulse.Utilities
{
    /// <summary>
    /// Converts a currency amount to English words — equivalent of the VB6 ConvertInWords function.
    /// </summary>
    public static class NumberToWordsConverter
    {
        private static readonly string[] Ones =
        {
            "", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE",
            "TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN",
            "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN"
        };

        private static readonly string[] Tens =
        {
            "", "", "TWENTY", "THIRTY", "FORTY", "FIFTY",
            "SIXTY", "SEVENTY", "EIGHTY", "NINETY"
        };

        /// <summary>
        /// Converts amount to words, e.g. 1234.56 USD → "USD ONE THOUSAND TWO HUNDRED THIRTY FOUR AND CENTS 56/100"
        /// </summary>
        public static string Convert(decimal amount, string currency = "USD")
        {
            if (amount < 0) return $"MINUS {Convert(-amount, currency)}";
            if (amount == 0) return $"{currency} ZERO ONLY";

            long wholePart = (long)Math.Truncate(amount);
            int centsPart = (int)Math.Round((amount - wholePart) * 100);

            string words = ConvertWholeNumber(wholePart);
            string result = $"{currency.ToUpper()} {words}";

            if (centsPart > 0)
                result += $" AND CENTS {centsPart:00}/100";
            else
                result += " ONLY";

            return result;
        }

        private static string ConvertWholeNumber(long n)
        {
            if (n == 0) return "ZERO";

            string words = "";

            if (n >= 1_000_000_000)
            {
                words += ConvertWholeNumber(n / 1_000_000_000) + " BILLION ";
                n %= 1_000_000_000;
            }
            if (n >= 1_000_000)
            {
                words += ConvertWholeNumber(n / 1_000_000) + " MILLION ";
                n %= 1_000_000;
            }
            if (n >= 1_000)
            {
                words += ConvertWholeNumber(n / 1_000) + " THOUSAND ";
                n %= 1_000;
            }
            if (n >= 100)
            {
                words += Ones[n / 100] + " HUNDRED ";
                n %= 100;
            }
            if (n >= 20)
            {
                words += Tens[n / 10];
                n %= 10;
                if (n > 0) words += " " + Ones[n];
            }
            else if (n > 0)
            {
                words += Ones[n];
            }

            return words.Trim();
        }
    }
}
