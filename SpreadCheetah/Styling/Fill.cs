using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadCheetah.Styling
{
    /// <summary>
    /// Represents the fill part of a <see cref="Style"/>.
    /// </summary>
    public sealed class Fill : IEquatable<Fill>
    {
        /// <summary>ARGB (alpha, red, green, blue) color of the fill.</summary>
        public Color? Color { get; set; }

        /// <inheritdoc/>
        public bool Equals(Fill? other) => other != null && EqualityComparer<Color?>.Default.Equals(Color, other.Color);

        /// <inheritdoc/>
        public override bool Equals(object? obj) => obj is Fill other && Equals(other);

        /// <inheritdoc/>
        public override int GetHashCode() => HashCode.Combine(Color);
    }
}
