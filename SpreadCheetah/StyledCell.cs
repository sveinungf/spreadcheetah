namespace SpreadCheetah
{
    public class StyledCell
    {
        public Cell Cell { get; }
        public StyleId? StyleId { get; }

        public StyledCell(Cell cell, StyleId? styleId)
        {
            Cell = cell;
            StyleId = styleId;
        }
    }
}
