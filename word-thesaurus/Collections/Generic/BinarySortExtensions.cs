using System;
using System.Collections.Generic;
using System.Linq;

namespace SeaRisenLib2.Collections.Generic
{
    public static class BinarySortExtensions
    {
        /// <summary>
        /// In-place sort of a list collection including WPF's ObservableCollection
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="comparer"></param>
        public static void BinarySort<T>(this IList<T> list, Func<T, T, int> comparer = null)
        {
            if (list.Count > 1) // a 0 or 1 length list doesn't need to be sorted
            {
                for (int i = 1; i < list.Count; i++)
                {
                    T key = list[i];
                    int index = Binary_SearchExtensions
                        .Binary_Search(list, key, comparer, 0, i);
                    if (index < 0)
                        index = ~index;
                    if (index != i)
                    {
                        list.RemoveAt(i);
                        //list.Remove(key);
                        list.Insert(index, key);
                    }
                }
            }
        }
    }
}
