using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadCheetah.Styling
{
    /// <summary>
    /// Represents the font part of a <see cref="Style"/>.
    /// </summary>
    public sealed class Font : IEquatable<Font>
    {
        /// <summary>Bold font weight. Defaults to <c>false</c>.</summary>
        public bool Bold { get; set; }

        /// <summary>Italic font type. Defaults to <c>false</c>.</summary>
        public bool Italic { get; set; }

        /// <summary>Adds a horizontal line through the center of the characters. Defaults to <c>false</c>.</summary>
        public bool Strikethrough { get; set; }

        /// <summary>ARGB (alpha, red, green, blue) color of the font.</summary>
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
