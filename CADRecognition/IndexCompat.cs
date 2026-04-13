namespace System
{
    public readonly struct Index
    {
        private readonly int _value;

        public Index(int value, bool fromEnd = false)
        {
            _value = fromEnd ? ~value : value;
        }

        public int Value => _value < 0 ? ~_value : _value;
        public bool IsFromEnd => _value < 0;

        public int GetOffset(int length)
        {
            var offset = _value;
            if (IsFromEnd)
            {
                offset += length + 1;
            }
            return offset;
        }

        public static Index Start => new Index(0);
        public static Index End => new Index(0, fromEnd: true);
        public static Index FromStart(int value) => new Index(value);
        public static Index FromEnd(int value) => new Index(value, fromEnd: true);
        public static implicit operator Index(int value) => new Index(value);

        public override string ToString()
        {
            return IsFromEnd ? "^" + Value.ToString() : Value.ToString();
        }
    }
}
