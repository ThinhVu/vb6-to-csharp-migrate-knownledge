using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;

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

    // This is my implemetation to using Strongly-typed Collection based-on VB6 Collection
    /// <summary>
    /// A Visual Basic Collection is an ordered set of items that can be referred
    /// to as a unit.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    [Serializable]
    [DebuggerDisplay("Count = {Count}")]
    public class VB6Collection<T> : IEnumerable<T>
    {
        Microsoft.VisualBasic.Collection mCol;

        /// <summary>
        /// Creates and returns a new Visual Basic Microsoft.VisualBasic.Collection object.
        /// </summary>
        public VB6Collection()
        {
            mCol = new Microsoft.VisualBasic.Collection();
        }

        /// <summary>
        /// Returns an Integer containing the number of elements in a collection. Read-only.
        /// </summary>       
        public int Count
        {
            get
            {
                return mCol.Count;
            }
        }

        /// <summary>
        /// Returns a specific element of a Collection object either by position or by
        /// key. Read-only.
        /// 
        /// </summary>
        /// <param name="index">
        /// (A) A numeric expression that specifies the position of an element of the
        /// collection. Index must be a number from 1 through the value of the collection's
        /// Microsoft.VisualBasic.Collection.Count property. Or (B) An Object expression
        /// that specifies the position or key string of an element of the collection.
        /// 
        /// </param>
        /// <returns>
        /// Returns a specific element of a Collection object either by position or by
        /// key. Read-only.
        /// </returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public T this[int index]
        {
            get
            {
                return (T)mCol[index];
            }
            // because VB Collection's item is immutable so setter is not neccessary.
        }

        /// <summary>
        /// Returns a specific element of a Collection object either by position or by
        /// key. Read-only.
        /// 
        /// </summary>
        /// <param name="key">
        /// A unique String expression that specifies a key string that can be used,
        /// instead of a positional index, to access an element of the collection. Key
        /// must correspond to the Key argument specified when the element was added
        /// to the collection.
        /// 
        /// </param>
        /// <returns>
        /// Returns a specific element of a Collection object either by position or by
        /// key. Read-only.
        /// </returns>
        public T this[string key]
        {
            get
            {
                return (T)mCol[key];
            }
            // because VB Collection's item is immutable so setter is not neccessary.
        }        

        /// <summary>
        /// Adds an element to a Collection object.
        /// 
        /// </summary>
        /// <param name="item">
        /// Required. An object of any type that specifies the element to add to the
        /// collection.
        /// </param>
        /// <param name="key">
        /// Optional. A unique String expression that specifies a key string that can
        /// be used instead of a positional index to access this new element in the collection.
        /// </param>
        /// <param name="before">
        /// Optional. An expression that specifies a relative position in the collection.
        /// The element to be added is placed in the collection before the element identified
        /// by the Before argument. If Before is a numeric expression, it must be a number
        /// from 1 through the value of the collection's Microsoft.VisualBasic.Collection.Count
        /// property. If Before is a String expression, it must correspond to the key
        /// string specified when the element being referred to was added to the collection.
        /// You cannot specify both Before and After.
        /// </param>
        /// <param name="after">
        /// Optional. An expression that specifies a relative position in the collection.
        /// The element to be added is placed in the collection after the element identified
        /// by the After argument. If After is a numeric expression, it must be a number
        /// from 1 through the value of the collection's Count property. If After is
        /// a String expression, it must correspond to the key string specified when
        /// the element referred to was added to the collection. You cannot specify both
        /// Before and After.        
        /// </param>
        public void Add(T item, string key = null, object before = null, object after = null)
        {
            mCol.Add(item, key, before, after);
        }

        /// <summary>
        /// Returns a Boolean value indicating whether a Visual Basic Collection object
        /// contains an element with a specific key.
        /// </summary>
        /// <param name="key">
        /// Required. A String expression that specifies the key for which to search 
        /// the elements of the collection.</param>
        /// <returns>
        /// Returns a Boolean value indicating whether a Visual Basic Collection object 
        /// contains an element with a specific key.
        /// </returns>
        public bool Contains(string key)
        {
            return mCol.Contains(key);
        }

        /// <summary>
        /// Returns a reference to an enumerator object, which is used to iterate over
        /// a Microsoft.VisualBasic.Collection object.
        /// </summary>
        /// <returns>
        /// Returns a reference to an enumerator object, which is used to iterate over
        /// a Microsoft.VisualBasic.Collection object.
        /// </returns>
        public IEnumerator<T> GetEnumerator()
        {
            foreach (var item in mCol)
            {
                yield return (T)item;
            }
        }

        /// <summary>
        /// Returns a reference to an enumerator object, which is used to iterate over
        /// a Microsoft.VisualBasic.Collection object.
        /// </summary>
        /// <returns>
        /// Returns a reference to an enumerator object, which is used to iterate over
        /// a Microsoft.VisualBasic.Collection object.
        /// </returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return mCol.GetEnumerator();
        }

        /// <summary>
        /// Removes an element from a Collection object.
        /// </summary>
        /// <param name="index">
        /// A numeric expression that specifies the position of an element of the collection.
        /// Index must be a number from 1 through the value of the collection's Microsoft.VisualBasic.Collection.Count
        /// property.
        /// </param>
        public void Remove(int index)
        {
            mCol.Remove(index);
        }

        /// <summary>
        /// Removes an element from a Collection object.
        /// </summary>
        /// <param name="key">
        /// A unique String expression that specifies a key string that can be used, 
        /// instead of a positional index, to access an element of the collection. Key 
        /// must correspond to the Key argument specified when the element was added 
        /// to the collection.</param>
        public void Remove(string key)
        {
            mCol.Remove(key);                        
        }

        /// <summary>
        /// Deletes all elements of a Visual Basic Collection object.
        /// </summary>
        public void Clear()
        {
            mCol.Clear();            
        }       
    }
}
