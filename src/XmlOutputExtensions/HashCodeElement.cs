namespace UsefulDataTools
{
    /// <summary>
    /// Internal class to help prevent infinite loops during recursive object searches.
    /// </summary>
    internal class HashCodeElement
    {
        private readonly int _currentHashCode;

        /// <summary>
        /// The <see cref="ParentHashCodeElement"/> will be set at every level in the search and can be compared using <see cref="HashCodeExistsInStructure"/>.
        /// </summary>
        public HashCodeElement ParentHashCodeElement { private get; set; }

        /// <summary>
        /// Create a new <see cref="HashCodeElement"/> using a generated <see cref="hashCode"/>.
        /// </summary>
        /// <param name="hashCode"></param>
        public HashCodeElement(int hashCode)
        {
            _currentHashCode = hashCode;
        }

        /// <summary>
        /// Check whether the <see cref="hashCode"/> already exists within the recursive structure.
        /// </summary>
        /// <param name="hashCode">A generated Hash Code which will be compared to all other hash codes in the structure.</param>
        /// <returns><see cref="bool"/></returns>
        public bool HashCodeExistsInStructure(int hashCode)
        {
            if (ParentHashCodeElement == null)
                return false;
            if (hashCode == ParentHashCodeElement._currentHashCode)
                return true;
            return ParentHashCodeElement.HashCodeExistsInStructure(hashCode);
        }
    }
}