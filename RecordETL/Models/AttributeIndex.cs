namespace RecordETL.Models
{
    public class AttributeIndex
    {
        public string Name { get; set; }

        private int _index = -1;
        public int Index { get => _index; set => _index = value; }
    }
}
