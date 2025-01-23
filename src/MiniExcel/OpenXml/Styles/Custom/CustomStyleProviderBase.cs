using MiniExcelLibs.OpenXml.Styles.Custom.Models;
using MiniExcelLibs.Utils;
using MiniExcelLibs.WriteAdapter;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs.OpenXml.Styles.Custom
{
    public abstract class CustomStyleProviderBase
    {
        public abstract int NumFmtCount { get; }

        public abstract int FontCount { get; }

        public abstract int FillCount { get; }

        public abstract int BorderCount { get; }

        public abstract int CellStyleXfCount { get; }

        public abstract int CellXfCount { get; }

        public abstract NumFmt GetNumFmt(int numFmtIndex);

        public virtual Task<NumFmt> GetNumFmtAsync(int numFmtIndex, CancellationToken cancellationToken = default) => Task.FromResult(GetNumFmt(numFmtIndex));

        public abstract Font GetFont(int fontIndex);

        public virtual Task<Font> GetFontAsync(int fontIndex, CancellationToken cancellationToken = default) => Task.FromResult(GetFont(fontIndex));

        public abstract Fill GetFill(int fillIndex);

        public virtual Task<Fill> GetFillAsync(int fillIndex, CancellationToken cancellationToken = default) => Task.FromResult(GetFill(fillIndex));

        public abstract Border GetBorder(int borderIndex);

        public virtual Task<Border> GetBorderAsync(int borderIndex, CancellationToken cancellationToken = default) => Task.FromResult(GetBorder(borderIndex));

        public abstract CellStyleXf GetCellStyleXf(int cellStyleXfIndex);

        public virtual Task<CellStyleXf> GetCellStyleXfAsync(int cellStyleXfIndex, CancellationToken cancellationToken = default) => Task.FromResult(GetCellStyleXf(cellStyleXfIndex));

        internal void CheckCellStyleXf(CellStyleXf cellStyleXf)
        {
            if (cellStyleXf.CustomNumberFormatIndex.HasValue && cellStyleXf.CustomNumberFormatIndex >= NumFmtCount)
            {
                throw new IndexOutOfRangeException($"Custom CellStyleXf.CustomNumberFormatIndex {cellStyleXf.CustomNumberFormatIndex} is out of range.");
            }
            if (cellStyleXf.CustomFontIndex >= FontCount)
            {
                throw new IndexOutOfRangeException($"Custom CellStyleXf.CustomFontIndex {cellStyleXf.CustomFontIndex} is out of range.");
            }
            if (cellStyleXf.CustomFillIndex >= FillCount)
            {
                throw new IndexOutOfRangeException($"Custom CellStyleXf.CustomFillIndex  {cellStyleXf.CustomFillIndex} is out of range.");
            }
            if (cellStyleXf.CustomBorderIndex >= BorderCount)
            {
                throw new IndexOutOfRangeException($"Custom CellStyleXf.CustomBorderIndex  {cellStyleXf.CustomBorderIndex} is out of range.");
            }
        }

        public abstract CellXf GetCellXf(int cellXfIndex);

        public virtual Task<CellXf> GetCellXfAsync(int cellXfIndex, CancellationToken cancellationToken = default) => Task.FromResult(GetCellXf(cellXfIndex));

        public abstract int? GetHeaderCellXfId(string sheetName, ExcelColumnInfo columnInfo);

        public abstract int? GetValueCellXfId(string sheetName, CellWriteInfo[] row, CellWriteInfo cellValue);
    }
}
