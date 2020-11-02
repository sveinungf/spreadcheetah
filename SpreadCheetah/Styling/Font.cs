using System;

namespace SpreadCheetah.Styling
{
    public sealed class Font : IEquatable<Font>
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }

        public bool Equals(Font? other) => other != null && Bold == other.Bold && Italic == other.Italic;
        public override bool Equals(object? obj) => obj is Font other && Equals(other);
        public override int GetHashCode() => HashCode.Combine(Bold, Italic);
    }
}
