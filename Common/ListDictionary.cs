using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
    /// <summary>
    /// What we should know about VB6 Collection class?
    /// - support key, value pair .like Dictionary.
    /// - manipulate element ( add, remove, insert ) like List.
    /// - immutable.
    /// - 1-based index.
    /// - object collection data structure.
    /// 
    /// Why VB6 Collection is bad?
    /// - Low performance.
    /// - Consume more memory than generic collection. (Because boxing)
    /// 
    /// 
    /// VB6 Collection is object collection data structure. Hence, when converting code, so much developers want to
    /// change this collection to Dictionary or List to gain better performance and familiar in use. But when you do that,
    /// please consider notice listed below:
    /// 
    /// 1) When you consider using Dictionary instead of Collection, you should care about :
    /// 1.1) whether VB6 code access item in Collection using index or just using foreach loop (enumerable) and doesn't care about index at all.
    ///     + VB6 Collection support access using index 1-based and Dictionary doesn't support it directly.
    ///     + Different in index base (0 vs 1) come up with so much problem.
    ///     
    /// 1.2) whether order of item added into Collection is important or not.
    ///     + VB6 Collection care about the order of added item. When you add new item to Collection without before or after provided,
    ///       this item always sit in the last position or "array".
    ///     + Dictionary doesn't care about the order of item, just 'fill' the blank space in "entries" array. "entries" array is private member of Dictionary which actually hold item values.
    ///     
    ///     E.g : Suppose when you remove item in Dictionary you made blank position in "entries" array, new added item will be placed at this position (fill the hole) so the different in C# Dictionary and VB6 occur.    
    ///     
    /// 2) When you want using List instead of Collection, you should care about :
    ///     + Whether VB6 code add item without key or not. (if not then you can using List instead of Collection but you still face with index based problem).
    ///     
    /// 3) Whether performance is critical.
    /// </summary>

    // This is my implemetation to using Strongly-typed Collection based-on List
    /// <summary>
    /// A Visual Basic Collection is an ordered set of items that can be referred
    /// to as a unit.
    /// </summary>    
    /// <typeparam name="V"></typeparam>
    [DebuggerDisplay("Count = {Count}")]
    public class ListDictionary<V> : IEnumerable<V>
    {
        private List<string> _keys;
        private List<V> _values;

        public ListDictionary()
        {
            _keys = new List<string>();
            _values = new List<V>();
        }

        public int Count
        {
            get
            {
                return _keys.Count;
            }
        }

        public V this[int index]
        {
            get
            {
                index--;
                if (index < 0 || index >= _keys.Count)
                    throw new IndexOutOfRangeException();
                return _values[index];
            }
        }
        public V this[string key]
        {
            get
            {
                ThrowIfKeyIsNull(key);
                ThrowIfKeyIsNotExist(key);
                return _values[_keys.IndexOf(key)];
            }
        }

        public List<string> Keys
        {
            get
            {
                // mutable in foreach
                return _keys.ToList();
            }
        }
        public List<V> Values
        {
            get
            {
                // mutable in foreach
                return _values.ToList();
            }
        }

        private void ThrowIfKeyIsNull(string key)
        {
            if (key == null)
                throw new ArgumentNullException("Key is null");
        }
        private void ThrowIfKeyIsNotExist(string key)
        {
            if (!ContainsKey(key))
                throw new ArgumentException("Key '" + key + "' not exist");
        }
        private void ThrowIfKeyExistAndNotNull(string key)
        {
            if (key != null && ContainsKey(key))
                throw new ArgumentException("Key '" + key + "' existed");
        }

        public void Add(string key, V value)
        {
            ThrowIfKeyExistAndNotNull(key);

            _keys.Add(key);
            _values.Add(value);
        }
        public void Add(string key, V value, object before = null, object after = null)
        {
            if (before != null)
            {
                switch (before.GetType().Name)
                {
                    case "String":
                        InsertBeforeKey(before.ToString(), key, value);
                        break;
                    case "Int32":
                        InsertBeforeIndex(Convert.ToInt32(before), key, value);
                        break;
                    default:
                        throw new ArgumentException("Type of before is not supported");
                }
            }
            else if (after != null)
            {
                switch (after.GetType().Name)
                {
                    case "String":
                        InsertAfterKey(after.ToString(), key, value);
                        break;
                    case "Int32":
                        InsertAfterIndex(Convert.ToInt32(after), key, value);
                        break;
                    default:
                        throw new ArgumentException("Type of after is not supported");
                }
            }
            else
            {
                Add(key, value);
            }
        }

        private void InsertBeforeKey(string before, string key, V value)
        {
            ThrowIfKeyIsNull(before);
            ThrowIfKeyIsNotExist(before);
            ThrowIfKeyExistAndNotNull(key);

            int iob = _keys.IndexOf(before);
            _keys.Insert(iob, key);
            _values.Insert(iob, value);
        }
        private void InsertBeforeIndex(int index, string key, V value)
        {
            // Insert before accept index == Count
            // When index == Count, insert is add new item to the end of List
            index--;
            if (index < 0 || index > Count)
                throw new IndexOutOfRangeException();

            ThrowIfKeyExistAndNotNull(key);

            _keys.Insert(index, key);
            _values.Insert(index, value);
        }
        private void InsertAfterKey(string after, string key, V value)
        {
            ThrowIfKeyIsNull(after);
            ThrowIfKeyIsNotExist(after);
            ThrowIfKeyExistAndNotNull(key);

            int index = _keys.IndexOf(after) + 1;
            _keys.Insert(index, key);
            _values.Insert(index, value);
        }
        private void InsertAfterIndex(int index, string key, V value)
        {
            index--;
            if (index < -1 || index >= Count)
                throw new IndexOutOfRangeException();

            ThrowIfKeyExistAndNotNull(key);

            index++;
            _keys.Insert(index, key);
            _values.Insert(index, value);
        }

        private bool ContainsKey(string key)
        {
            return _keys.IndexOf(key) != -1;
        }
        private bool ContainsValue(V value)
        {
            return _values.IndexOf(value) != -1;
        }

        public void Clear()
        {
            _keys.Clear();
            _values.Clear();
        }
        public bool Remove(int index)
        {
            index--;
            if (index < 0 || index > Count)
                throw new IndexOutOfRangeException();

            _keys.RemoveAt(index);
            _values.RemoveAt(index);
            return true;
        }
        public bool Remove(string key)
        {
            var iok = _keys.IndexOf(key);
            if (iok != -1)
            {
                _keys.RemoveAt(iok);
                _values.RemoveAt(iok);
                return true;
            }
            return false;
        }

        public bool TryGetValue(string key, out V value)
        {
            value = default(V);
            var iok = _keys.IndexOf(key);
            if (iok != -1)
            {
                value = _values[iok];
                return true;
            }
            return false;
        }

        public IEnumerator<V> GetEnumerator()
        {
            // clone enumerator to allow edit collection in foreach
            return Values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            // clone enumerator to allow edit collection in foreach
            return Values.GetEnumerator();
        }
    }
}
