namespace RequirementsCompiler.Maps
{
    public class StyleDefinitions
    {        
        public static IStyle Heading1
        {
            get
            {
                return new Style("Heading1", "Heading1")
                    {
                        Bold = true,
                        Color = "FF0000",
                        FontName = "Arial",
                        FontSize = 25,
                        MarginBottom = 10,
                        MarginTop = 10
                    };
            }
        }

        public static IStyle Heading2
        {
            get
            {
                return new Style("Heading2", "Heading2")
                {
                    Bold = false,
                    Color = "0000FF",
                    FontName = "Arial",
                    FontSize = 20,
                    MarginTop = 20,
                    MarginBottom = 0
                };
            }
        }
    }
}