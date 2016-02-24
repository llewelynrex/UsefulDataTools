namespace UsefulDataTools
{
    internal class HashCodeElement
    {
        private int CurrentHashCode { get; }

        public HashCodeElement ParentHashCodeElement { get; set; }

        public HashCodeElement(int hashCode)
        {
            CurrentHashCode = hashCode;
        }

        public bool HashKeyExistsInStructure(int hashKey)
        {
            if (ParentHashCodeElement == null)
                return false;
            if (hashKey == ParentHashCodeElement.CurrentHashCode)
                return true;
            return ParentHashCodeElement.HashKeyExistsInStructure(hashKey);
        }
    }
}