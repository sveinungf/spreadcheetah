using System;
using System.Collections.Generic;

namespace SpreadCheetah.Styling
{
    public sealed class Style : IEquatable<Style>
    {
        public Font Font { get; set; } = new Font();

        public bool Equals(Style? other)
        {
            return other != null && EqualityComparer<Font>.Default.Equals(Font, other.Font);
        }

        public override bool Equals(object? obj)
        {
            return obj is Style other && Equals(other);
        }

        public override int GetHashCode()
        {
            return 1417255472 + EqualityComparer<Font>.Default.GetHashCode(Font);
        }
    }
}
