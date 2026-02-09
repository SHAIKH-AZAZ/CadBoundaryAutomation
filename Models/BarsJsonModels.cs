using System;
using System.Collections.Generic;

namespace CadBoundaryAutomation.Models
{
    // ✅ Grouped bars (no coordinates)
    public class BarJson
    {
        public string Orientation { get; set; }   // "Horizontal" / "Vertical"
        public double Length { get; set; }        // already rounded to 2 decimals in BarGenerator
        public int Repetition { get; set; }       // how many bars have same rounded length
        public List<string> Handles { get; set; } // handles of all those bars
    }

    public class BarsRunJson
    {
        public string DrawingName { get; set; }
        public string DrawingPath { get; set; }
        public DateTime CreatedAt { get; set; }

        public string BoundaryHandle { get; set; }

        public string Orientation { get; set; }
        public double SpacingH { get; set; }
        public double SpacingV { get; set; }

        public int TotalBars { get; set; }
        public double TotalLength { get; set; }

        public List<BarJson> Bars { get; set; }
    }
}
