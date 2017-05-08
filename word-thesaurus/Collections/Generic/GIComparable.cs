using System;
using System.Reflection;

namespace SeaRisenLib2.Collections.Generic
{
    public static class GIComparable
    {
        /// <summary>
        /// Get Compare method from T if it derives from IComparable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static Func<T, T, int> Compare<T>() // where T : IComparable
        {
            if (typeof(IComparable).IsAssignableFrom(typeof(T)))
            {
                // Get IComparable of T if it is of interface IComparable
                var staticMethod = typeof(T).GetMethod("Compare",
                    BindingFlags.Static | BindingFlags.Public,
                    null,
                    new Type[] { typeof(T), typeof(T) },
                    null);
                return (Func<T, T, int>)Delegate.CreateDelegate(typeof(Func<T, T, int>), staticMethod);
            }
            return null;
        }
    }
}
