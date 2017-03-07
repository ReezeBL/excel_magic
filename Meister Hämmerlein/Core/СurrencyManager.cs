using System;
using System.Collections.Generic;
using System.Xml;

namespace Meister_Hämmerlein.Core
{
    public static class СurrencyManager
    {
        private static readonly Dictionary<DateTime, decimal> Rates = new Dictionary<DateTime, decimal>();

        public static decimal GetRate(DateTime date)
        {
            decimal rate;
            if (!Rates.TryGetValue(date, out rate))
            {
                var usdXml = GetUsdXml(date);
                var xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(usdXml);

                var usdRate = xmlDocument.SelectSingleNode("Valute/Value")?.InnerText;
                if (usdRate != null)
                    rate = Rates[date] = decimal.Parse(usdRate);
                else
                    rate = 1;

                Console.WriteLine($"Loaded rate on {date}, it is: {rate}");
            }
            return rate;
        }

        private static string GetUsdXml(DateTime date)
        {
            var reader = new XmlTextReader($"http://www.cbr.ru/scripts/XML_daily.asp?date_req={date.Day:D2}/{date.Month:D2}/{date.Year:D4}");
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || !reader.HasAttributes) continue;

                while (reader.MoveToNextAttribute())
                {
                    if (reader.Name != "ID" || reader.Value != "R01235") continue;
                    reader.MoveToElement();
                    return reader.ReadOuterXml();
                }
            }
            return "";
        }
    }
}
