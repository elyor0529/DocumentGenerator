namespace RequirementsCompiler.Maps
{
    public class Style : IStyle
    {
        public Style(string name, string id)
        {
            this.Name = name;
            this.Id = id;
        }

        public string Name { get; private set; }

        public string Id { get; private set;  }

        public string FontName { get; set; }

        public string Color { get; set; }

        public bool Bold { get; set; }

        public int FontSize { get; set; }

        public string BasedOn { get; set; }

        public int MarginTop { get; set; }

        public int MarginBottom { get; set; }
    }
}