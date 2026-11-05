using System;
using System.IO;
using System.Net;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Collections.Generic;
using System.Globalization;

using ExcelTools.Tables;

namespace ExcelTools.TableDiff {
    internal static class HtmlDiffWriter {
        private static string Escape(string text) {
            return WebUtility.HtmlEncode(text ?? "");
        }

        private static string Asset(string name) {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name)) {
                if (stream == null) {
                    throw new InvalidOperationException("Missing embedded report asset: " + name);
                }
                using (var reader = new StreamReader(stream, Encoding.UTF8)) {
                    return reader.ReadToEnd();
                }
            }
        }

        private static string Content(string text) {
            if (text.StartsWith("=", StringComparison.Ordinal) || text.Length > 350 || text.Count(c => c == '\n') > 3) {
                var first = text.Split('\n')[0];
                var shortText = first.Substring(0, Math.Min(100, first.Length));
                return "<details class=\"formula\"><summary>" + Escape(shortText) + "…</summary><pre>" + Escape(text) + "</pre></details>";
            }
            return "<pre>" + Escape(text) + "</pre>";
        }

        public static string Generate(Configuration configuration, List<SheetDiff> differences) {
            var sections = new StringBuilder();
            var table = new StringBuilder("<table><tr><th>Старый лист</th><th>Новый лист</th><th>Ключ</th><th>Порог</th><th>Изменено</th><th>Добавлено</th><th>Удалено</th><th>Было строк</th><th>Стало строк</th></tr>");
            var nav = new StringBuilder();
            for (var index = 0; index < differences.Count; index++) {
                var sheetDiff = differences[index];
                var id = "s" + index;
                table.Append("<tr><td>").Append(Escape(sheetDiff.OldSheet)).Append("</td><td>")
                    .Append(Escape(sheetDiff.NewSheet)).Append("</td><td>").Append(Escape(sheetDiff.Key))
                    .Append("</td><td>").Append((sheetDiff.Threshold * 100).ToString("0.##", CultureInfo.InvariantCulture)).Append("%</td>");
                foreach (var count in new[] { sheetDiff.Count("change"), sheetDiff.Count("add"), sheetDiff.Count("delete"), sheetDiff.OldCount, sheetDiff.NewCount }) {
                    table.Append("<td>").Append(count).Append("</td>");
                }
                table.Append("</tr>");
                nav.Append("<a href=\"#").Append(id).Append("\">").Append(Escape(sheetDiff.Name)).Append("</a>");
                sections.Append("<section id=\"").Append(id).Append("\"><h2>").Append(Escape(sheetDiff.Name))
                    .Append(" <small>~").Append(sheetDiff.Count("change")).Append(" / +").Append(sheetDiff.Count("add"))
                    .Append(" / −").Append(sheetDiff.Count("delete")).Append("</small></h2>");
                sections.Append("<p class=\"muted\">Ключ: ").Append(Escape(sheetDiff.Key)).Append(" · Порог: ")
                    .Append((sheetDiff.Threshold * 100).ToString("0.##", CultureInfo.InvariantCulture)).Append("%</p>");
                if (sheetDiff.Changes.Count == 0) {
                    sections.Append("<p class=\"muted\">Изменений данных и формул нет.</p>");
                }
                foreach (var change in sheetDiff.Changes) {
                    AppendChange(sections, sheetDiff, change, sheetDiff.Key);
                }
                sections.Append("</section>");
            }
            table.Append("</table>");
            return "<!doctype html><html lang=\"ru\"><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">"
                + "<title>Сравнение Excel</title><style>" + Asset("report.css") + "</style><main><h1>Excel · изменения между версиями</h1>"
                + "<p class=\"muted\">Профиль: " + Escape(configuration.Profile) + "</p>"
                + "<p class=\"muted\">" + Escape(Path.GetFileName(configuration.Old)) + " → " + Escape(Path.GetFileName(configuration.New))
                + " · − старая версия / + новая версия</p><div class=\"stats\"><b>~ " + differences.Sum(s => s.Count("change"))
                + " изменено</b><b>+ " + differences.Sum(s => s.Count("add")) + " добавлено</b><b>− "
                + differences.Sum(s => s.Count("delete")) + " удалено</b></div>"
                + "<div class=\"legend\"><span class=\"green\">Добавлено значение</span><span class=\"purple\">Заменено значение · внутри показано «Было»</span><span class=\"red\">Удалено значение</span></div>"
                + "<details><summary>Сводка и способ сравнения</summary>" + table
                + "<p>Сравниваются сохранённые значения и тексты формул только в указанных парах листов. Читаются все строки, включая скрытые и исключённые фильтрами Excel; сохранённые фильтры не ограничивают сравнение. Ключ и порог каждой пары приведены в сводке. Оформление и кэшированные результаты вычислений не сравниваются. Пустые строки и пустые оформленные столбцы не считаются записями. При разных заголовках распознанной таблицы сравнение останавливается.</p>"
                + "<p>Границы колонок определяются по непрерывным непустым заголовкам слева и справа от ключевого поля. Колонки за этими границами исключены. Диапазоны определяются независимо в обеих версиях, поэтому горизонтальный сдвиг таблицы не считается изменением.</p>"
                + "<p>Сначала сопоставляются полностью одинаковые строки независимо от позиции с учётом повторов. Остальные записи сопоставляются по ключу своей пары и наибольшей доле совпадающих заполненных полей, не ниже заданного для этой пары порога. Равные по близости кандидаты отмечены для проверки. Строки вне таблиц сопоставляются по номеру. Перестановки одинаковых записей не показаны как изменения.</p></details><nav>"
                + nav + "<input placeholder=\"Поиск по ФИО, полю или значению\"><select><option value=\"\">Все изменения</option><option value=\"change\">Изменённые</option><option value=\"add\">Добавленные</option><option value=\"delete\">Удалённые</option></select></nav>"
                + sections + "</main><script>" + Asset("report.js") + "</script></html>";
        }

        private static void AppendChange(StringBuilder output, SheetDiff sheetDiff, RowChange rowChange, string key) {
            var header = sheetDiff.NewHeader ?? sheetDiff.OldHeader;
            var row = rowChange.After ?? rowChange.Before;
            var columnOffset = rowChange.After == null ? sheetDiff.OldColumnOffset : sheetDiff.NewColumnOffset;
            var keyColumn = sheetDiff.OldHeader == null ? -1 : Array.IndexOf(sheetDiff.OldHeader.Values, key);
            var identity = keyColumn < 0 ? sheetDiff.Name : row.ValueAt(keyColumn);
            var labels = Enumerable.Range(0, sheetDiff.Width)
                .Select(c => header == null || header.ValueAt(c).Length == 0 ? CellAddress.ColumnName(columnOffset + c + 1) : header.ValueAt(c)).ToArray();
            output.Append("<article data-kind=\"").Append(rowChange.Kind).Append("\"><header><b>").Append(Escape(identity))
                .Append("</b><span class=\"row-context\">");
            output.Append(String.Join(" · ", Enumerable.Range(0, sheetDiff.Width)
                .Where(c => labels[c] == "Подія" && row.ValueAt(c).Length > 0).Select(c => Escape(row.ValueAt(c)))));
            output.Append("</span><span class=\"row-location\">")
                .Append(rowChange.Before == null ? "—" : rowChange.Before.Number.ToString())
                .Append(" → ").Append(rowChange.After == null ? "—" : rowChange.After.Number.ToString()).Append("</span></header>");
            if (rowChange.Ambiguous) {
                output.Append("<p class=\"warning\">Несколько записей имеют одинаковую близость: проверьте сопоставление.</p>");
            }
            output.Append("<div class=\"row-scroll\"><div class=\"row-cells\">");
            for (var column = 0; column < sheetDiff.Width; column++) {
                if (header != null && header.ValueAt(column).Length == 0) {
                    continue;
                }
                var x = rowChange.Before == null ? "" : rowChange.Before.ValueAt(column);
                var y = rowChange.After == null ? "" : rowChange.After.ValueAt(column);
                var unchanged = (rowChange.Kind == "change" && x == y) || (x.Length == 0 && y.Length == 0);
                var status = unchanged ? "unchanged" : x.Length > 0 && y.Length > 0 ? "replacement" : y.Length > 0 ? "insertion" : "removal";
                var tag = unchanged ? "Без изменений" : status == "replacement" ? "Заменено" : status == "insertion" ? "Добавлено" : "Удалено";
                output.Append("<div class=\"cell ").Append(status).Append("\"><div class=\"cell-heading\"><b>")
                    .Append(Escape(labels[column])).Append("</b><span>").Append(CellAddress.ColumnName(columnOffset + column + 1)).Append(row.Number)
                    .Append("</span></div><span class=\"cell-tag\">").Append(tag).Append("</span>");
                if (status == "replacement") {
                    output.Append("<div class=\"previous\"><span>Было · ").Append(CellAddress.ColumnName(sheetDiff.OldColumnOffset + column + 1))
                        .Append(rowChange.Before.Number).Append("</span>").Append(Content(x)).Append("</div>");
                }
                var value = y.Length > 0 ? y : x;
                output.Append("<div class=\"current\">").Append(unchanged ? "<pre>-- // --</pre>" : Content(value)).Append("</div>");
                if (unchanged) {
                    output.Append("<template class=\"original-value\">")
                        .Append(value.Length > 0 ? Content(value) : "<pre>Пустая ячейка</pre>").Append("</template>");
                }
                output.Append("</div>");
            }
            output.Append("</div></div></article>");
        }
    }
}