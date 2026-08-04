using System.Collections.Generic;
using System.Linq;

namespace Fuwoa.Core.ExportCount
{
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
    /// 输入字符串数组，输出按计数降序排列的唯一值列表。
    /// </summary>
    public class ExportCountService
    {
        /// <summary>
        /// 对输入数据进行分组计数，按计数降序排列。
        /// </summary>
        /// <param name="items">输入的字符串数组</param>
        /// <returns>按计数降序排列的结果列表</returns>
        public List<CountResultItem> Compute(string[] items)
        {
            if (items == null || items.Length == 0)
                return new List<CountResultItem>();

            return items
                .GroupBy(x => x)
                .Select(g => new CountResultItem { Value = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Count)
                .ToList();
        }
    }
}
