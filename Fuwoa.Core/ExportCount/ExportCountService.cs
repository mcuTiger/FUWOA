using System;
using System.Collections.Generic;
using System.Linq;

namespace Fuwoa.Core.ExportCount
{
    /// <summary>
    /// 排序方式。
    /// </summary>
    public enum SortMode
    {
        ByCount,
        ByTitle
    }

    /// <summary>
    /// 导出计数结果项。
    /// </summary>
    public class CountResultItem
    {
        public string Value { get; set; }
        public int Count { get; set; }
    }

    /// <summary>
    /// 导出计数核心服务。
    /// </summary>
    public class ExportCountService
    {
        public List<CountResultItem> Compute(string[] items, SortMode sortMode = SortMode.ByCount,
            bool descending = true)
        {
            if (items == null || items.Length == 0)
                return new List<CountResultItem>();

            var groups = items
                .GroupBy(x => x)
                .Select(g => new CountResultItem { Value = g.Key, Count = g.Count() });

            var cmp = StringComparer.CurrentCulture;
            if (sortMode == SortMode.ByTitle)
                return descending
                    ? groups.OrderByDescending(x => x.Value, cmp).ThenByDescending(x => x.Count).ToList()
                    : groups.OrderBy(x => x.Value, cmp).ThenBy(x => x.Count).ToList();
            else
                return descending
                    ? groups.OrderByDescending(x => x.Count).ThenBy(x => x.Value, cmp).ToList()
                    : groups.OrderBy(x => x.Count).ThenBy(x => x.Value, cmp).ToList();
        }
    }
}
