using System;
using System.Collections;
using System.Collections.Generic;

namespace SeaRisenLib2.Collections.Generic
{
    /// <summary>
    /// Convert Func T, T, int to IComparer T.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class LambdaCompare<T> : IComparer<T>
    {
        readonly Func<T, T, int> comparer;

        public LambdaCompare(Func<T, T, int> f)
        {
            comparer = f;
        }

        public int Compare(T x, T y)
        {
            return comparer(x, y);
        }
    }

    public class LambdaCompare : IComparer
    {
        readonly Func<object, object, int> comparer;

        public LambdaCompare(Func<object, object, int> f) { comparer = f; }

        public int Compare(object x, object y)
        {
            return comparer(x, y);
        }
    }
}
