using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Чтение листа «link». В нём по строке на каждый магазин для каждого АЦР - десятки
/// тысяч строк, поэтому читаются только строки нужных АЦР, а сведения по ним сразу
/// сворачиваются в одну запись.
/// </summary>
public static class LinkSheetReader
{
    /// <summary>
    /// Заголовки «link». «Концепт» необязателен: по нему узнаются острова, у которых
    /// «Группа_ам» пустая. Без него острова узнаются только по грейду O120 / O140.
    /// </summary>
    public static HeaderMap ResolveHeaders(SheetGrid grid)
    {
        var headers = HeaderResolver.Resolve(grid, PriceSchema.LinkSheet, PriceSchema.Link.Specs);
        var key = TextUtils.NormalizeKey(PriceSchema.Link.Concept);

        for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
        {
            if (string.Equals(TextUtils.NormalizeKey(grid.Text(headers.HeaderRow, column)), key, StringComparison.Ordinal))
            {
                var columns = new Dictionary<string, int>(headers.Columns, StringComparer.Ordinal)
                {
                    [PriceSchema.Link.Concept] = column,
                };
                return new HeaderMap(headers.HeaderRow, columns);
            }
        }

        return headers;
    }

    public static void Read(
        SheetGrid grid,
        HeaderMap headers,
        IReadOnlySet<string> wantedAcr,
        Dictionary<string, List<LinkRow>> result)
    {
        var acrColumn = headers[PriceSchema.Link.Acr];
        var hasConcept = headers.TryGet(PriceSchema.Link.Concept, out var conceptColumn);
        var firstRow = Math.Max(headers.HeaderRow + 1, grid.FirstRow);

        for (var row = firstRow; row <= grid.LastRow; row++)
        {
            var acr = TextUtils.NormalizeKey(grid.Text(row, acrColumn));
            if (acr.Length == 0 || !wantedAcr.Contains(acr))
            {
                continue;
            }

            if (!result.TryGetValue(acr, out var rows))
            {
                rows = new List<LinkRow>();
                result[acr] = rows;
            }

            rows.Add(new LinkRow(
                acr,
                grid.Number(row, headers[PriceSchema.Link.DenyGoodsDivision]),
                grid.Number(row, headers[PriceSchema.Link.DenyGoodsAll]),
                grid.Number(row, headers[PriceSchema.Link.Included]),
                TextUtils.Normalize(grid.Text(row, headers[PriceSchema.Link.StoreGrade])),
                grid.Number(row, headers[PriceSchema.Link.DateStart]),
                grid.Number(row, headers[PriceSchema.Link.DateEnd]),
                hasConcept ? TextUtils.Normalize(grid.Text(row, conceptColumn)) : string.Empty));
        }
    }

    /// <summary>
    /// Сворачивает строки одного АЦР в запись: по каждому грейду магазинов («Группа_ам») -
    /// сколько магазинов и скольким из них АЦР положен («Вкл» = 1 и обе «Deny goods» = 0),
    /// и даты.
    ///
    /// Так грейд проверяет и инструкция: «Группа_ам» даёт грейд магазина, а «Вкл»
    /// и «Deny goods» - положен ли ему АЦР. Колонка «Группы_ок» одинакова во всех строках
    /// и о запретах конкретных грейдов ничего не говорит.
    /// </summary>
    public static LinkEntry Summarize(string acr, IReadOnlyList<LinkRow> rows)
    {
        // Острова без грейда («Группа_ам» пустая, «Концепт» - остров) собираются в свою группу:
        // по ним сверяется, отмечены ли острова в «Цены», а не пишется «магазины без грейда».
        var grades = rows
            .GroupBy(GroupName, StringComparer.Ordinal)
            .Select(group => new LinkGrade(
                group.Key,
                group.Count(),
                group.Count(row => row.IsAllowed)))
            .OrderBy(grade => GradeKey(grade.Grade), StringComparer.Ordinal)
            .ToList();

        // Даты сверяются с «Периодом продаж магазины», поэтому сначала берутся даты магазинов,
        // а у островов - только если у магазинов дат нет.
        var dated = rows.OrderBy(row => IsIsland(row) ? 1 : 0).ToList();

        return new LinkEntry(
            acr,
            grades,
            dated.Select(r => r.DateStart).FirstOrDefault(d => d.HasValue),
            dated.Select(r => r.DateEnd).FirstOrDefault(d => d.HasValue),
            rows.Count,
            rows.Count(row => row.IsAllowed));
    }

    private static string GroupName(LinkRow row)
    {
        var grade = TextUtils.Normalize(row.StoreGrade);
        if (grade.Length > 0)
        {
            // Одно имя на грейд, как бы его ни набрали: «О120» и «O120» - одна группа.
            return PriceSchema.Prices.AllGrades.FirstOrDefault(known =>
                       string.Equals(GradeKey(known), GradeKey(grade), StringComparison.Ordinal))
                   ?? grade;
        }

        return IsIsland(row) ? PriceSchema.Link.IslandsWithoutGrade : string.Empty;
    }

    /// <summary>Строка острова: грейд O120 / O140 либо, при пустом грейде, «Концепт» - остров.</summary>
    private static bool IsIsland(LinkRow row) =>
        PriceSchema.Prices.IslandGrades.Any(island =>
            string.Equals(GradeKey(row.StoreGrade), GradeKey(island), StringComparison.Ordinal)) ||
        (TextUtils.Normalize(row.StoreGrade).Length == 0 &&
         TextUtils.NormalizeKey(row.Concept).Contains(PriceSchema.Link.IslandConceptMarker, StringComparison.Ordinal));

    /// <summary>
    /// Ключ грейда для сравнения: без регистра и без разницы между латиницей и кириллицей -
    /// «О120» и «O120», «С1» и «C1» набирают по-разному, а грейд один.
    /// </summary>
    public static string GradeKey(string? grade)
    {
        const string cyrillic = "АВСЕНКМОРТХ";
        const string latin = "ABCEHKMOPTX";

        var chars = TextUtils.Normalize(grade).Replace(" ", string.Empty).ToUpperInvariant().ToCharArray();
        for (var i = 0; i < chars.Length; i++)
        {
            var index = cyrillic.IndexOf(chars[i]);
            if (index >= 0)
            {
                chars[i] = latin[index];
            }
        }

        return new string(chars);
    }
}
