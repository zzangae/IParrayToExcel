using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;

namespace IpToolsLib
{
    [Guid("A3A27F6F-7D61-4C0E-9A22-1D0D3E7A1F10")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IIpTools
    {
        [DispId(1)]
        object[,] Sort(object values, bool ascending = true, bool deduplicate = false);

        [DispId(2)]
        bool IsValid(string ip);

        [DispId(3)]
        string Normalize(string ip);

        [DispId(4)]
        object[,] Unique(object values);

        [DispId(5)]
        string ExtractIp(string text);
    }

    [Guid("D5E0AA6B-9A0F-4A59-99DA-7E2E4F9E5B33")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("IpTools.IpTools")]
    public class IpTools : IIpTools
    {
        private static IEnumerable<string> FlattenToStrings(object values)
        {
            if (values == null)
                yield break;

            Array arr = values as Array;
            if (arr != null)
            {
                int rank = arr.Rank;
                if (rank == 1)
                {
                    int lower = arr.GetLowerBound(0);
                    int upper = arr.GetUpperBound(0);
                    for (int i = lower; i <= upper; i++)
                    {
                        object v = arr.GetValue(i);
                        if (v != null) yield return v.ToString().Trim();
                    }
                }
                else if (rank == 2)
                {
                    int rLower = arr.GetLowerBound(0);
                    int rUpper = arr.GetUpperBound(0);
                    int cLower = arr.GetLowerBound(1);
                    int cUpper = arr.GetUpperBound(1);

                    for (int r = rLower; r <= rUpper; r++)
                    {
                        for (int c = cLower; c <= cUpper; c++)
                        {
                            object v = arr.GetValue(r, c);
                            if (v != null) yield return v.ToString().Trim();
                        }
                    }
                }
                else
                {
                    foreach (object v in arr)
                    {
                        if (v != null) yield return v.ToString().Trim();
                    }
                }
            }
            else
            {
                yield return values.ToString().Trim();
            }
        }

        public object[,] Sort(object values, bool ascending = true, bool deduplicate = false)
        {
            List<string> items = FlattenToStrings(values)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToList();

            if (deduplicate)
            {
                HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                List<string> distinctItems = new List<string>();
                foreach (string s in items)
                {
                    if (seen.Add(s))
                        distinctItems.Add(s);
                }
                items = distinctItems;
            }

            List<KeyValuePair<string, ulong?>> list = new List<KeyValuePair<string, ulong?>>();
            foreach (string s in items)
            {
                ulong key;
                bool ok = TryMakeSortKey(s, out key);
                list.Add(new KeyValuePair<string, ulong?>(s, ok ? (ulong?)key : null));
            }

            list.Sort(delegate (KeyValuePair<string, ulong?> a, KeyValuePair<string, ulong?> b)
            {
                if (a.Value.HasValue && !b.Value.HasValue)
                    return -1;
                if (!a.Value.HasValue && b.Value.HasValue)
                    return 1;
                if (a.Value.HasValue && b.Value.HasValue)
                    return a.Value.Value.CompareTo(b.Value.Value);
                return string.Compare(a.Key, b.Key, StringComparison.OrdinalIgnoreCase);
            });

            if (!ascending)
                list.Reverse();

            string[] result = list.Select(p => p.Key).ToArray();
            return ToColumn(result);
        }

        public bool IsValid(string ip)
        {
            ulong dummy;
            return TryMakeSortKey(ip, out dummy);
        }

        public string Normalize(string ip)
        {
            IPAddress addr;
            if (!IPAddress.TryParse(ip, out addr))
                return string.Empty;
            return addr.ToString();
        }

        public object[,] Unique(object values)
        {
            HashSet<string> set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<string> result = new List<string>();
            foreach (string s in FlattenToStrings(values))
            {
                if (!string.IsNullOrWhiteSpace(s) && set.Add(s))
                    result.Add(s);
            }
            return ToColumn(result.ToArray());
        }

        public string ExtractIp(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;

            string[] parts = text.Split(new[] { ' ', '\t', ',', ';', '[', ']', '(', ')', '{', '}', '<', '>', '|', '=' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string p in parts)
            {
                IPAddress addr;
                if (IPAddress.TryParse(p.Trim(), out addr))
                    return p.Trim();
            }
            return string.Empty;
        }

        private static bool TryMakeSortKey(string ip, out ulong key)
        {
            key = 0;
            IPAddress addr;
            if (!IPAddress.TryParse(ip, out addr))
                return false;

            byte[] bytes = addr.GetAddressBytes();
            if (addr.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork && bytes.Length == 4)
            {
                key = ((ulong)bytes[0] << 24) | ((ulong)bytes[1] << 16) | ((ulong)bytes[2] << 8) | bytes[3];
                return true;
            }

            return false;
        }

        private static object[,] ToColumn(string[] items)
        {
            object[,] result = new object[items.Length, 1];
            for (int i = 0; i < items.Length; i++)
                result[i, 0] = items[i];
            return result;
        }
    }
}
