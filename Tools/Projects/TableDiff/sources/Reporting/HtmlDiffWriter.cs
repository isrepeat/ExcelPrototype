using System;
using System.IO;
using System.Net;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelTools.TableDiff {
    public static class HtmlDiffWriter {
        private static string Escape(string text) {
            return WebUtility.HtmlEncode(text ?? "");
        }

        private static string TabFileName(string path) {
            var name = Path.GetFileNameWithoutExtension(path);
            var characters = StringInfo.ParseCombiningCharacters(name);
            return characters.Length > TableDiffConstants.TabFileNameMaxLength
                ? name.Substring(0, characters[TableDiffConstants.TabFileNameMaxLength - 3]) + "…"
                : name;
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

        public static void Write(string output, string profile, string oldFile, string newFile, List<ExcelTools.Changes.SheetDiff> differences, string method) {
            Directory.CreateDirectory(Path.GetDirectoryName(output));
            var temporary = ExcelTools.Tooling.OutputPublisher.CreateTemporaryPath(output, ".html");
            try {
                File.WriteAllText(temporary, Generate(profile, oldFile, newFile, differences, method), new UTF8Encoding(false));
                ExcelTools.Tooling.OutputPublisher.Publish(temporary, output);
            }
            finally {
                if (File.Exists(temporary)) {
                    File.Delete(temporary);
                }
            }
        }

        private static string Generate(string profile, string oldFile, string newFile, List<ExcelTools.Changes.SheetDiff> differences, string method) {
            var sections = new StringBuilder();
            var payloads = new StringBuilder();
            var summary = new StringBuilder("<table class=\"summary-table\"><thead><tr><th>Старый лист</th><th>Новый лист</th><th>Ключ</th><th>Порог</th><th>Изменено</th><th>Добавлено</th><th>Удалено</th><th>Было строк</th><th>Стало строк</th></tr></thead><tbody>");
            for (var index = 0; index < differences.Count; index++) {
                var sheetDiff = differences[index];
                summary.Append("<tr><td>").Append(Escape(sheetDiff.OldSheet)).Append("</td><td>")
                    .Append(Escape(sheetDiff.NewSheet)).Append("</td><td>").Append(Escape(sheetDiff.Key))
                    .Append("</td><td>").Append(Percent(sheetDiff.Threshold)).Append("</td>");
                foreach (var count in new[] { sheetDiff.Count("change"), sheetDiff.Count("add"), sheetDiff.Count("delete"), sheetDiff.OldCount, sheetDiff.NewCount }) {
                    summary.Append("<td>").Append(count).Append("</td>");
                }
                summary.Append("</tr>");
                AppendSheet(sections, payloads, sheetDiff, "s" + index);
            }
            summary.Append("</tbody></table>");
            return "<!doctype html><html lang=\"ru\"><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">"
                + "<title>" + Escape(TabFileName(oldFile)) + " → " + Escape(TabFileName(newFile))
                + "</title><style>" + Asset("report.css") + "</style>"
                + "<header class=\"report-header\"><h1>Сравнение Excel</h1><span class=\"profile\">" + Escape(profile)
                + "</span><div class=\"stats\" aria-label=\"Итоги сравнения\"><span>~ " + differences.Sum(s => s.Count("change"))
                + " изменено</span><span>+ " + differences.Sum(s => s.Count("add"))
                + " добавлено</span><span>− " + differences.Sum(s => s.Count("delete")) + " удалено</span></div></header>"
                + "<p class=\"source-files\">" + Escape(Path.GetFileName(oldFile)) + " → " + Escape(Path.GetFileName(newFile)) + "</p>"
                + "<details class=\"report-summary\"><summary>Сводка и способ сравнения</summary><div class=\"summary-scroll\">" + summary + "</div>"
                + "<p>Сравниваются сохранённые значения и тексты формул только в указанных парах листов. Скрытые и отфильтрованные строки включены; оформление и кэшированные результаты формул не сравниваются.</p>"
                + (method != null ? "<p>" + Escape(method) + "</p>" :
                "<p>Границы колонок определяются по непрерывным непустым заголовкам вокруг ключевого поля независимо в каждой версии. Колонки сопоставляются по точным названиям: + означает новую колонку, − удалённую. Порог близости рассчитывается только по общим колонкам.</p>"
                + "<p>Сначала исключаются точные совпадения с учётом повторов. Остальные записи сопоставляются по ключу и порогу своей пары; равные по близости кандидаты отмечаются знаком ?. Строки вне таблиц сопоставляются по номеру.</p>")
                + "<p>Старое значение зачёркнуто, новое расположено ниже. Каждое значение ограничено тремя видимыми строками с многоточием независимо от другого значения. Поиск учитывает полный текст. Неизменённые формулы можно показать переключателем в панели поиска. Адреса ячеек доступны во всплывающей подсказке.</p></details>"
                + "<nav class=\"report-toolbar\" aria-label=\"Поиск и фильтр\"><div class=\"toolbar-controls\" id=\"toolbar-controls\"><input type=\"search\" aria-label=\"Поиск по значению или изменённому полю\" placeholder=\"Поиск по значению или полю\">"
                + "<select aria-label=\"Тип изменений\"><option value=\"\">Все изменения</option><option value=\"change\">Изменённые</option><option value=\"add\">Добавленные</option><option value=\"delete\">Удалённые</option></select>"
                + "<label class=\"formula-toggle\"><input type=\"checkbox\" id=\"show-unchanged-formulas\">Неизменённые формулы</label>"
                + "<label class=\"formula-toggle\"><input type=\"checkbox\" id=\"live-column-resize\" checked>Изменять ширину сразу</label>"
                + "<label class=\"formula-toggle\"><input type=\"checkbox\" id=\"hide-unchanged-columns\">Скрывать неизменённые колонки</label>"
                + "<span class=\"legend\"><span class=\"legend-add\">+ добавлено</span><span class=\"legend-change\">~ заменено</span><span class=\"legend-delete\">− удалено</span></span></div>"
                + "<div class=\"table-navigation\" role=\"group\" aria-label=\"Навигация по таблице\">"
                + "<label class=\"column-scale\">Ширина <input type=\"range\" id=\"column-scale\" aria-label=\"Масштаб ширины колонок\" min=\"55\" max=\"190\" step=\"15\" value=\"100\"><output id=\"column-scale-value\" for=\"column-scale\">100%</output></label>"
                + "<button type=\"button\" id=\"table-top\" title=\"К началу таблицы\" aria-label=\"К началу таблицы\">↑</button>"
                + "<button type=\"button\" id=\"table-bottom\" title=\"К концу таблицы\" aria-label=\"К концу таблицы\">↓</button></div></nav>"
                + "<main>" + sections + "</main>" + payloads + "<noscript>Для просмотра отчёта включите JavaScript.</noscript><script>" + Asset("report.js") + "</script></html>";
        }

        private static string Percent(double value) {
            return (value * 100).ToString("0.##", CultureInfo.InvariantCulture) + "%";
        }

        private static void AppendSheet(StringBuilder output, StringBuilder payloads, ExcelTools.Changes.SheetDiff sheetDiff, string id) {
            var header = sheetDiff.NewHeader ?? sheetDiff.OldHeader;
            var offset = sheetDiff.NewHeader == null ? sheetDiff.OldColumnOffset : sheetDiff.NewColumnOffset;
            var columns = Enumerable.Range(0, sheetDiff.Width)
                .Where(c => header == null || header.ValueAt(c).Length > 0
                    || (sheetDiff.OldHeader != null && sheetDiff.OldHeader.ValueAt(c).Length > 0)).ToList();
            output.Append("<section id=\"").Append(id).Append("\" data-sheet-name=\"").Append(Escape(sheetDiff.Name))
                .Append("\"><div class=\"sheet-heading\"><h2>").Append(Escape(sheetDiff.Name)).Append("</h2><span>Ключ: ")
                .Append(Escape(sheetDiff.Key)).Append(" · ").Append(Percent(sheetDiff.Threshold))
                .Append(" · ~").Append(sheetDiff.Count("change")).Append(" / +").Append(sheetDiff.Count("add"))
                .Append(" / −").Append(sheetDiff.Count("delete")).Append("</span></div>")
                .Append("<nav class=\"pagination\" aria-label=\"Страницы различий\">")
                .Append("<button type=\"button\" data-page-action=\"first\" aria-label=\"Первая страница\">«</button>")
                .Append("<button type=\"button\" data-page-action=\"previous\" aria-label=\"Предыдущая страница\">‹</button>")
                .Append("<span class=\"page-status\" role=\"status\" aria-live=\"polite\"></span>")
                .Append("<button type=\"button\" data-page-action=\"next\" aria-label=\"Следующая страница\">›</button>")
                .Append("<button type=\"button\" data-page-action=\"last\" aria-label=\"Последняя страница\">»</button></nav>")
                .Append("<div class=\"table-host\"></div><p class=\"empty-message\" hidden></p></section>");

            // Intern cell strings once per sheet instead of repeating full HTML for every cell.
            var strings = new List<string> { "" };
            var indices = new Dictionary<string, int>(StringComparer.Ordinal) { { "", 0 } };
            var rows = new StringBuilder("[");
            var first = true;
            foreach (var change in sheetDiff.Changes) {
                if (!first) {
                    rows.Append(',');
                }
                first = false;
                rows.Append('[').Append(change.Kind == "change" ? 0 : change.Kind == "add" ? 1 : 2)
                    .Append(',').Append(change.Before == null ? 0 : change.Before.Number)
                    .Append(',').Append(change.After == null ? 0 : change.After.Number)
                    .Append(',').Append(change.Ambiguous ? 1 : 0).Append(',');
                AppendValues(rows, change.Before, columns, strings, indices);
                rows.Append(',');
                AppendValues(rows, change.After, columns, strings, indices);
                if (change.ReviewNote != null) {
                    rows.Append(',');
                    AppendJsonString(rows, change.ReviewNote);
                }
                rows.Append(']');
            }
            rows.Append(']');
            payloads.Append("<script type=\"application/json\" id=\"data-").Append(id).Append("\">{\"columns\":[");
            for (var index = 0; index < columns.Count; index++) {
                if (index > 0) {
                    payloads.Append(',');
                }
                var column = columns[index];
                payloads.Append('[').Append(column).Append(',');
                var name = header == null ? ExcelTools.Tables.CellAddress.ColumnName(offset + column + 1) : header.ValueAt(column);
                if (name.Length == 0 && sheetDiff.OldHeader != null) {
                    name = sheetDiff.OldHeader.ValueAt(column);
                }
                AppendJsonString(payloads, name);
                var oldColumn = sheetDiff.OldColumns == null ? sheetDiff.OldColumnOffset + column : sheetDiff.OldColumns[column];
                var newColumn = sheetDiff.NewColumns == null ? sheetDiff.NewColumnOffset + column : sheetDiff.NewColumns[column];
                payloads.Append(',').Append(oldColumn).Append(',').Append(newColumn);
                payloads.Append(']');
            }
            payloads.Append("],\"offsets\":[").Append(sheetDiff.OldColumnOffset).Append(',').Append(sheetDiff.NewColumnOffset)
                .Append("],\"strings\":[");
            for (var index = 0; index < strings.Count; index++) {
                if (index > 0) {
                    payloads.Append(',');
                }
                AppendJsonString(payloads, strings[index]);
            }
            payloads.Append("],\"rows\":").Append(rows).Append("}</script>");
        }

        private static void AppendValues(
            StringBuilder output,
            ExcelTools.Tables.RowSnapshot rowSnapshot,
            List<int> columns,
            List<string> strings,
            Dictionary<string, int> indices) {
            if (rowSnapshot == null) {
                output.Append("null");
                return;
            }
            output.Append('[');
            for (var index = 0; index < columns.Count; index++) {
                if (index > 0) {
                    output.Append(',');
                }
                var value = rowSnapshot.ValueAt(columns[index]);
                int stringIndex;
                if (!indices.TryGetValue(value, out stringIndex)) {
                    stringIndex = strings.Count;
                    indices.Add(value, stringIndex);
                    strings.Add(value);
                }
                output.Append(stringIndex);
            }
            output.Append(']');
        }

        private static void AppendJsonString(StringBuilder output, string value) {
            output.Append('"');
            foreach (var character in value) {
                if (character == '"' || character == '\\') {
                    output.Append('\\').Append(character);
                }
                else if (character < 32 || character == '<' || character == '>' || character == '&'
                    || character == '\u2028' || character == '\u2029') {
                    // Escape script delimiters as well as control characters without changing the data.
                    output.Append("\\u").Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                }
                else {
                    output.Append(character);
                }
            }
            output.Append('"');
        }
    }
}