using System;
using System.Collections.Generic;

namespace SpreadCheetah.Styling
{
    /// <summary>
    /// Represents style for one or more worksheet cells.
    /// </summary>
    public sealed class Style : IEquatable<Style>
    {
        /// <summary>Fill for the cell.</summary>
        public Fill Fill { get; set; } = new Fill();

        /// <summary>Font for the cell's value.</summary>
        public Font Font { get; set; } = new Font();

        /// <inheritdoc/>
        public bool Equals(Style? other) => other != null
            && EqualityComparer<Fill>.Default.Equals(Fill, other.Fill)
            && EqualityComparer<Font>.Default.Equals(Font, other.Font);

        /// <inheritdoc/>
        public override bool Equals(object? obj) => obj is Style other && Equals(other);

        /// <inheritdoc/>
        public override int GetHashCode() => HashCode.Combine(Fill, Font);
    }
}
