using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadCheetah.Styling
{
    public sealed class Font : IEquatable<Font>
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Strikethrough { get; set; }
        public Color? Color { get; set; }

        /// <inheritdoc/>
        public bool Equals(Font? other) => other != null
            && Bold == other.Bold && Italic == other.Italic && Strikethrough == other.Strikethrough
            && EqualityComparer<Color?>.Default.Equals(Color, other.Color);

        /// <inheritdoc/>
        public override bool Equals(object? obj) => obj is Font other && Equals(other);

        /// <inheritdoc/>
        public override int GetHashCode() => HashCode.Combine(Bold, Italic, Strikethrough, Color);
    }
}
