using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadCheetah.Styling
{
    public sealed class Fill : IEquatable<Fill>
    {
        public Color? Color { get; set; }

        /// <inheritdoc/>
        public bool Equals(Fill? other) => other != null && EqualityComparer<Color?>.Default.Equals(Color, other.Color);

        /// <inheritdoc/>
        public override bool Equals(object? obj) => obj is Fill other && Equals(other);

        /// <inheritdoc/>
        public override int GetHashCode() => HashCode.Combine(Color);
    }
}
