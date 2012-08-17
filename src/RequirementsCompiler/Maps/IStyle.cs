namespace RequirementsCompiler.Maps
{
    public interface IStyle
    {
        string Name { get; }

        string Id { get; }

        string FontName { get; }

        string Color { get; }

        bool Bold { get; }

        int FontSize { get;  }

        string BasedOn { get; }

        int MarginTop { get; }

        int MarginBottom { get; }
    }
}
