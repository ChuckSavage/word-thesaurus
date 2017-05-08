using System.Collections.Generic;

namespace SeaRisenLib2.Collections
{
    public interface IOrderedList<T> : IList<T>
    {
        int Add(T item);
        void Clear();
        bool Contains(T item);
        void CopyTo(T[] array, int arrayIndex);
        int Remove(T item);
        int Count { get; }
        bool IsReadOnly { get; }
        //IEnumerator GetEnumerator();
        IEnumerator<T> GetEnumerator();
        int IndexOf(T item);
        void Insert(int index, T item);
        void RemoveAt(int index);
        T this[int index] { get; set; }
    }
}
