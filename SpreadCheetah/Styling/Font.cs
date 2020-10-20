using System;

namespace SpreadCheetah.Styling
{
    public sealed class Font : IEquatable<Font>
    {
        public bool Bold { get; set; }

        public bool Equals(Font? other)
        {
            return other != null && Bold == other.Bold;
        }

        public override bool Equals(object? obj)
        {
            return obj is Font other && Equals(other);
        }

        public override int GetHashCode()
        {
            return 768441094 + Bold.GetHashCode();
        }
    }
}
