using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace SeaRisenLib2.Collections.Generic
{
    /// <summary>
    /// Extensions to the list.BinarySearch() method
    /// </summary>
    public static class BinarySearchExtensions
    {
        public static int Compare<T>(T a, T b, Func<T, T, int> comparer = null)
        {
            if (null != comparer)
                return comparer(a, b);
            if (a is IComparable)
                return ((IComparable)a).CompareTo(b);
            if (a is IComparer)
                return ((IComparer)a).Compare(a, b);
            return Comparer<T>.Default.Compare(a, b);
        }

        /// <summary>
        /// Perform a Binary Search on the collection, returning the
        /// index of the item.
        /// <para>Uses the type's IComparable comparer.</para>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        public static int BinarySearch<T>
        (
            this IEnumerable<T> source,
            T item
        )
        {
            return BinarySearch(source, item, null);
        }

        /// <summary>
        /// Perform a Binary Search on the collection, returning the
        /// index of the item.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <param name="item"></param>
        /// <param name="comparer">Can be null and will use the default Comparer for the type.</param>
        /// <returns></returns>
        public static int BinarySearch<T>
        (
            this IEnumerable<T> source,
            T item,
            Func<T, T, int> comparer
        )
        {
            List<T> list = source as List<T>;
            if (null == list)
                list = source.ToList();
            if (null == comparer)
                return list.BinarySearch(item, Comparer<T>.Default);
            return list
                .BinarySearch(item, new LambdaCompare<T>(comparer));
        }

        public static bool Contains_BinarySearch<TSource>
        (
            this IEnumerable<TSource> source,
            TSource value
        )
        {
            return Contains_BinarySearch(source, value, Comparer<TSource>.Default);
        }

        public static bool Contains_BinarySearch<TSource>
        (
            this IEnumerable<TSource> source,
            TSource value,
            IComparer<TSource> comparer
        )
        {
            return IndexOf_BinarySearch(source, value, comparer) >= 0;
        }

        public static int IndexOf_BinarySearch<TSource>
        (
            this IEnumerable<TSource> source,
            TSource value
        )
        {
            return IndexOf_BinarySearch(source, value, Comparer<TSource>.Default);
        }

        public static int IndexOf_BinarySearch<TSource>
        (
            this IEnumerable<TSource> source,
            TSource value,
            IComparer<TSource> comparer
        )
        {
            return Array.BinarySearch<TSource>(source.ToArray(), value, comparer);
        }
    }
}
