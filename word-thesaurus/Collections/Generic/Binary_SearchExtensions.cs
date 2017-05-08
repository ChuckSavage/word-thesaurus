using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace SeaRisenLib2.Collections.Generic
{
    // Iterative BinarySearch method (from: http://en.wikipedia.org/wiki/Binary_search_algorithm)
    /// <summary>
    /// Custom Binary Search Extensions
    /// <para>Where the normal list.BinarySearch() method isn't sufficient</para>
    /// </summary>
    public static class Binary_SearchExtensions
    {
        public static int Binary_Search<T>(
            this IList<T> list,
            T key = default(T),
            Func<T, T, int> comparer = null,
            int start = 0, int count = -1)
        {
            if (count < 0 || count > list.Count) count = list.Count;
            if (count == 0)
                return ~start;
            int mid = 0, min = start, max = count, midIs = 0;
            if (min < 0) min = 0;
            else if (min > max - 1) min = max - 1;

            try
            {
                // continue searching while [min,max] is not empty
                while (max >= min)
                {
                    // calculate the midpoint for roughly equal partition
                    mid = midpoint(min, max);

                    // determine which sub-array to search
                    if (mid >= count) // out-of-range
                        return ~mid;
                    midIs = BinarySearchExtensions.Compare(list[mid], key, comparer);
                    if (midIs < 0) // change min index to search upper sub-array
                        min = mid + 1;
                    else if (midIs > 0) // change max index to search lower sub-array
                        max = mid - 1;
                    else // key found at index mid
                        return mid;
                }
                // key not found
                if (midIs < 0)
                    return ~min;
                return ~mid;
            }
            catch (ArgumentOutOfRangeException ex)
            {
                throw new Exception(
                    string.Format(
                        "Binary_Search: start({0}), count({1}), min({2}), max({3}), mid({4})",
                        start, count, min, max, mid),
                    ex);
            }
        }

        public static int Binary_Search<T>
        (
            this IEnumerable<T> source,
            T key,
            Func<T, T, int> comparer,
            int min = 0, int max = -1
        )
        {
            return Binary_Search(source.ToList(), key, comparer, min, max);
        }

        /// <summary>
        /// Perform a custom Binary Search on the collection, returning the
        /// index of the item.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="comparer"></param>
        /// <param name="start"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        public static int Binary_Search<T>
        (
            this IList<T> list,
            Func<T, int> comparer,
            int start = 0, int count = -1
        )
        {
            Func<T, T, int> comp = (a, k) => comparer(a);
            return Binary_Search(list, default(T), comp, start, count);
        }

        /// <summary>
        /// Perform a custom Binary Search on the collection, returning the
        /// index of the item.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="S"></typeparam>
        /// <param name="source"></param>
        /// <param name="item"></param>
        /// <param name="comparer"></param>
        /// <returns></returns>
        public static int Binary_Search<T>
        (
            this IEnumerable<T> source,
            Func<T, int> comparer,
            int min = 0, int max = -1
        )
        {
            return Binary_Search(source.ToList(), comparer, min, max);
        }

        static int midpoint(int imin, int imax) { return (imax - imin) / 2 + imin; }

        /// <summary>
        /// Binary Search for item, deferring equality for the end.  Best used for
        /// searches where there are duplicate items, so that it returns the first
        /// duplicate found.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <param name="comparer">Comparer needs to check for less than key item and equal</param>
        /// <returns></returns>
        public static int Binary_Search_Deferred<T>
        (
            this IEnumerable<T> source,
            Func<T, int> comparer
        )
        {
            List<T> list = source as List<T>;
            if (null == list)
                list = source.ToList();
            if (null == comparer)
                throw new ArgumentNullException("Comparer function cannot be null.");

            if (list.Count <= 0)
                return ~0;

            // Iterative BinarySearch method (from: http://en.wikipedia.org/wiki/Binary_search_algorithm)
            int mid = 0, min = 0, max = list.Count;

            // inclusive indices
            //   0 <= min when using truncate toward zero divide
            //     mid = (min+max)/2;
            //   min unrestricted when using truncate toward minus infinity divide
            //     mid = (min+max)>>1; or
            //     mid = (int)floor((min+max)/2.0);

            // continually narrow search until just one element remains
            while (min < max)
            {
                mid = midpoint(min, max);

                // code must guarantee the interval is reduced at each iteration
                Debug.Assert(mid < max);
                // note: 0 <= min < max implies mid will always be less than max

                // reduce the search
                if (comparer(list[mid]) < 0)        //if (A[mid] < key)
                    min = mid + 1;
                else
                    max = mid;
            }
            // At exit of while:
            //   if A[] is empty, then max < min
            //   otherwise max == min

            if (min == list.Count)
                return ~min;
            // deferred test for equality
            if ((max == min) && (0 == comparer(list[min])))
                return min;
            else
                return ~min;
        }
    }
}
