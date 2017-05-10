using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using SeaRisenLib2.Collections.Generic;

namespace SeaRisenLib2.Collections
{
	/// <summary>
	/// <para>An Ordered List.</para>
	/// <para>If the Generic Class inherits IComparable, then comparer to constructor can be null.</para>
	/// <para>Thread safe.</para>
	/// </summary>
	/// <typeparam name="T"></typeparam>
	[DebuggerDisplay("Count = {Count}")]
	public class OrderedList<T> : IOrderedList<T>, IList
	{
		protected readonly object lockob = new object();
		public bool PermitDuplicates { get; private set; }
		/// <summary>
		/// Replace on add if duplicate and duplicates are not allowed.
		/// </summary>
		public bool Replace { get; private set; }

		/// <summary>
		/// Set new Comparer to list.  
		/// <para>WARNING: Setting this re-sorts the list.</para>
		/// </summary>
		public Func<T, T, int> Comparer
		{
			get { return _Comparer; }
			set
			{
				_Comparer = value;
				Sort();
			}
		}
		Func<T, T, int> _Comparer;

		readonly List<T> list = new List<T>();
		/* Test case
        static OrderedList()
        {
            List<int> list = Enumerable.Range(1, 10).ToList();
            //list.AddRange(Enumerable.Range(1, 10));
            //var shuffled = list.ShuffleAll(new Random());
            list.Reverse();
            var shuffled = list;
            //OrderedList<int> ordered = new OrderedList<int>(shuffled, (a, b) => a - b, false);
            OrderedList<int> ordered = new OrderedList<int>(shuffled, (a, b) => b - a, false);
            var test = ordered;
        }
        */

		/// <summary>
		/// Create OrderedList for a type that implents IComparable.
		/// All base types like string, int and float all implement IComparable.
		/// 
		/// <remarks>
		/// Permits no duplicates and 
		/// doesn't replace a duplicate if found.
		/// </remarks>
		/// </summary>
		public OrderedList()
			: this(null, false, false)
		{
			if (null == Comparer)
				throw new NotImplementedException("Generic Type must implement IComparable to use this constructor.");
		}

		/// <summary>
		/// Create OrderedList based on parameters
		/// </summary>
		/// <param name="comparer">If null and T is IComparable, will use T.Comparer() as the comparer</param>
		/// <param name="permitDuplicates">The list can contain duplicate items.</param>
		/// <param name="replace">If permitDuplicates is false, and replace true, new items with the same compared value of 0, will be replaced.</param>
		public OrderedList(Func<T, T, int> comparer = null, bool permitDuplicates = true, bool replace = true)
		{
			if (null == comparer)
				comparer = GIComparable.Compare<T>();
			Comparer = comparer;
			PermitDuplicates = permitDuplicates;
			if (!PermitDuplicates)
				Replace = replace;
		}

		public OrderedList
		(
			IEnumerable<T> range,
			Func<T, T, int> comparer,
			bool permitDuplicates = true,
			bool replace = true,
			bool rangeIsAlreadySorted = false
		)
			: this(comparer, permitDuplicates, replace)
		{
			if (rangeIsAlreadySorted)
				list.AddRange(range);
			else
				Add(range);
		}

		/// <summary>
		/// Adds item to list.
		/// </summary>
		/// <param name="item"></param>
		/// <returns>Index of insertion (positive value if duplicates permitted)</returns>
		public virtual int Add(T item)
		{
			if (null == item)
				throw new ArgumentNullException("Item argument to method cannot be null.");
			int index = IndexOf(item);
			lock (lockob)
			{
				if (index < 0)
					list.Insert(~index, item);
				else if (PermitDuplicates)
					list.Insert(index, item);
				else if (Replace)
				{
					T old = list[index];
					list.Insert(index, item);
					list.Remove(old);
				}
				return index;
			}
		}

		/// <summary>
		/// Add an array of items
		/// </summary>
		/// <param name="items"></param>
		public virtual void Add(params T[] items)
		{
			if (null == items)
				throw new ArgumentNullException("Items argument to Add() cannot be null");
			foreach (T name in items)
				Add(name);
		}

		/// <summary>
		/// Add an enumerable of items
		/// </summary>
		/// <param name="items"></param>
		public virtual void Add(IEnumerable<T> items)
		{
			if (null == items)
				throw new ArgumentNullException("Items argument to Add() cannot be null");
			Add(items.ToArray());
		}

		/// <summary>
		/// Add already sorted items to the end of the ordered list
		/// </summary>
		/// <param name="items"></param>
		/// <exception cref="ArgumentOutOfRangeException" />
		public virtual void AppendSorted(IEnumerable<T> items)
		{
			if (Count > 0)
			{
				int compare = Comparer(list.Last(), items.First());
				if (compare > 0 || (compare >= 0 && !PermitDuplicates))
					throw new ArgumentOutOfRangeException("AppendSorted: Items out of order");
			}
			if (default(T) == null && items.Any(i => i == null))
				throw new ArgumentNullException("Items argument contains a null.");
			lock (lockob)
				list.AddRange(items);
		}

		/// <summary>
		/// Clear the List
		/// </summary>
		public virtual void Clear() { lock (lockob) list.Clear(); }

		/// <summary>
		/// Make a copy of this OrderedList and pass it back full of the current items.
		/// </summary>
		/// <returns></returns>
		public OrderedList<T> Clone()
		{
			lock (lockob)
			{
				var clone = new OrderedList<T>(Comparer, PermitDuplicates, Replace);
				clone.list.AddRange(this.list);
				return clone;
			}
		}

		public virtual int Count { get { lock (lockob) return list.Count; } }

		/// <summary>
		/// Loop over each item in OrderedList.
		/// <remarks>Not thread safe.</remarks>
		/// </summary>
		/// <param name="action"></param>
		public void ForEach(Action<T> action)
		{
			var list = this.list.ToList(); // If someone wants to use Remove in the loop, let them
			foreach (T item in list)
				action(item);
		}

		/// <summary>
		/// Loop over each item in OrderedList.
		/// <remarks>Not thread safe.</remarks>
		/// </summary>
		/// <param name="action"></param>
		public void ForEach(Action<T, int> action)
		{
			var list = this.list.ToList(); // If someone wants to use RemoveAt in the loop, let them
			for (int index = 0; index < list.Count; index++)
				action(list[index], index);
		}

		/// <summary>
		/// Find Index of item or its opposite where to insert the item if not found.
		/// </summary>
		public virtual int IndexOf(T item)
		{
			if (null == item)
				throw new ArgumentNullException("Item argument to method cannot be null.");
			if (null == Comparer)
				throw new ArgumentNullException("Comparer cannot be null.");
			lock (lockob)
				return list.BinarySearch(item, Comparer);
		}

		/// <summary>
		/// Find Index of item or its opposite where to insert the item if not found.
		/// </summary>
		public virtual int IndexOf(Func<T, int> comparer)
		{
			if (null == comparer)
				throw new ArgumentNullException("OrderedList<T>.IndexOf().Comparer cannot be null.");
			lock (lockob)
			{
				if (PermitDuplicates)
					return list.Binary_Search_Deferred(comparer);
				return list.Binary_Search(comparer);
			}
		}

		/// <summary>
		/// Find Item if found or default(T) if not.
		/// </summary>
		public virtual T Item(Func<T, int> comparer)
		{
			int index = IndexOf(comparer);
			lock (lockob)
				if (index >= 0)
					return list[index];
			return default(T);
		}

		/// <summary>
		/// Find Item if found or default(T) if not.
		/// <para>Generic Type must be of type IComparable to use this method.</para>
		/// </summary>
		/// <param name="key">value to search for in list</param>
		/// <returns></returns>
		public virtual T Item(object key)
		{
			return Item(i => ((IComparable)i).CompareTo(key));
		}

		/// <summary>
		/// Load a (new) list of already sorted items
		/// </summary>
		/// <param name="items"></param>
		public virtual void LoadSorted(IEnumerable<T> items)
		{
			Clear();
			lock (lockob)
				list.AddRange(items);
		}

		/// <summary>
		/// Remove an array of items.
		/// </summary>
		/// <param name="items"></param>
		public virtual void Remove(params T[] items)
		{
			if (null == items)
				throw new ArgumentNullException("Items argument to Add() cannot be null");
			foreach (T name in items)
				Remove(name);
		}

		/// <summary>
		/// Remove an enumerable of items
		/// </summary>
		/// <param name="items"></param>
		public virtual void Remove(IEnumerable<T> items)
		{
			if (null == items)
				throw new ArgumentNullException("Items argument to Add() cannot be null");
			Remove(items.ToArray());
		}

		/// <summary>
		/// This will remove the correct item if this is a non-repeating list or
		/// the comparer can determine the correct item in a repeating list.
		/// </summary>
		/// <param name="item"></param>
		/// <returns></returns>
		public virtual int Remove(T item)
		{
			if (null == item)
				throw new ArgumentNullException("Item argument to method cannot be null.");
			int index = IndexOf(item);
			RemoveAt(index);
			return index;
		}

		/// <summary>
		/// Remove item and return its index if it was found.
		/// <para>Generic Type must be of type IComparable to use this method.</para>
		/// </summary>
		/// <param name="key">value to search for in list</param>
		/// <returns></returns>
		public virtual int Remove(object key)
		{
			T item;
			return Remove(key, out item);
		}

		/// <summary>
		/// Remove item and return its index if it was found.
		/// <para>Generic Type must be of type IComparable to use this method.</para>
		/// </summary>
		/// <param name="key">value to search for in list</param>
		/// <param name="item"></param>
		/// <returns></returns>
		public virtual int Remove(object key, out T item)
		{
			return Remove(i => ((IComparable)i).CompareTo(key), out item);
		}

		public virtual int Remove(Func<T, int> comparer)
		{
			T item;
			return Remove(comparer, out item);
		}

		public virtual int Remove(Func<T, int> comparer, out T item)
		{
			int index = IndexOf(comparer);
			RemoveAt(index, out item);
			return index;
		}

		public virtual void RemoveAt(int index)
		{
			T item;
			RemoveAt(index, out item);
		}

		public virtual void RemoveAt(int index, out T item)
		{
			lock (lockob)
				if (index >= 0 && index < Count)
				{
					item = this[index];
					list.RemoveAt(index);
				}
				else
					item = default(T);
		}

		/// <summary>
		/// Reverses the list.
		/// <para>WARNING: Comparer may be invalid if adding new items or calling IndexOf(item).</para>
		/// </summary>
		public virtual void Reverse()
		{
			lock (lockob)
				list.Reverse();
		}

		/// <summary>
		/// Re-sort the list.
		/// </summary>
		public virtual void Sort()
		{
			lock (lockob)
				list.BinarySort(Comparer);
			/*
            if (Count > 0)
            {
                Stopwatch a = Stopwatch.StartNew();
                var copy = new List<T>(list);
                list.Clear();
                Add(copy);
                // Binary Sort is much faster
                a.Stop();
                Stopwatch b = Stopwatch.StartNew();
                copy.BinarySort(Comparer);
                b.Stop();
                var at = a.ElapsedMilliseconds;
                var bt = b.ElapsedMilliseconds;
                at = bt;
            }
             */
		}

		/// <summary>
		/// Re-sort the list
		/// </summary>
		public virtual void Sort(Action<T, int> itemAdded)
		{
			lock (lockob)
				if (Count > 0)
				{
					var copy = new List<T>(list);
					list.Clear();
					foreach (T item in copy)
					{
						int index = Add(item);
						itemAdded(item, index);
					}
				}
		}

		public IEnumerator<T> GetEnumerator()
		{
			return list.GetEnumerator();
		}

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return list.GetEnumerator();
		}

		/// <summary>
		/// Insert item into ordered list.
		/// Does up to two compares to verify that it is being inserted properly.
		/// </summary>
		public virtual void Insert(int index, T item)
		{
			lock (lockob)
			{
				bool greaterThanPrevious = index == 0 || Comparer(list[index - 1], item) < 0;
				bool lessThanNext = index >= Count || Comparer(item, list[index]) < 0;
				if (greaterThanPrevious && lessThanNext)
					list.Insert(index, item);
				else
					throw new InvalidInsertException("Index for item puts item out of order", greaterThanPrevious, lessThanNext, index);
			}
		}

		[System.Runtime.CompilerServices.IndexerName("Items")]
		public virtual T this[int index]
		{
			get { lock (lockob) return list[index]; }
			set
			{
				lock (lockob)
				{
					bool greaterThanPrevious = index == 0 || Comparer(list[index - 1], value) < 0;
					bool lessThanNext = index != Count - 1 || Comparer(value, list[index + 1]) < 0;
					if (greaterThanPrevious && lessThanNext)
						list[index] = value;
					else
						throw new InvalidInsertException("Index for item puts item out of order", greaterThanPrevious, lessThanNext, index);
				}
			}
		}

		public class InvalidInsertException : Exception
		{
			public InvalidInsertException(string message, bool gtp, bool ltn, int index)
				: base(message)
			{
				GreaterThanPrevious = gtp;
				LessThanNext = ltn;
				Index = index;
			}
			public bool GreaterThanPrevious { get; private set; }
			public bool LessThanNext { get; private set; }
			public int Index { get; private set; }
		}

		void ICollection<T>.Add(T item)
		{
			Add(item);
		}

		public bool Contains(T item)
		{
			return IndexOf(item) >= 0;
		}

		public bool Contains(Func<T, int> comparer)
		{
			return IndexOf(comparer) >= 0;
		}

		/// <summary>
		/// Not Implemented
		/// </summary>
		public void CopyTo(T[] array, int arrayIndex)
		{
			// apparently called by Linq.ToList() methods
			list.CopyTo(array, arrayIndex);
		}

		public bool IsReadOnly
		{
			get { return false; }
		}

		bool ICollection<T>.Remove(T item)
		{
			return Remove(item) >= 0;
		}

		#region IList

		#endregion

		public int Add(object value)
		{
			return Add((T)value);
		}

		public bool Contains(object value)
		{
			return Contains((T)value);
		}

		public int IndexOf(object value)
		{
			return IndexOf((T)value);
		}

		public void Insert(int index, object value)
		{
			Insert(index, (T)value);
		}

		public bool IsFixedSize
		{
			get { return false; }
		}

		void IList.Remove(object value)
		{
			Remove((T)value);
		}

		object IList.this[int index]
		{
			get
			{
				return this[index];
			}
			set
			{
				this[index] = (T)value;
			}
		}

		public void CopyTo(Array array, int index)
		{
			lock (lockob)
				foreach (var item in list)
					array.SetValue(item, index++);
		}

		public bool IsSynchronized
		{
			get { return false; }
		}

		public object SyncRoot
		{
			get { throw new NotImplementedException(); }
		}

		/// <summary>
		/// New OrderedList of type T, no duplicates and doesn't replace if found on add.
		/// </summary>
		public static OrderedList<O> NewOrderedList<O>
		(
			IEnumerable<O> items = null,
			bool itemsSorted = false
		) where O : IComparable
		{
			Func<O, O, int> compare = (a, b) => ((IComparable)a).CompareTo(b);
			if (null == items)
				return new OrderedList<O>(compare, false, false);
			return new OrderedList<O>(items, compare, false, false, itemsSorted);
		}
	}
}
