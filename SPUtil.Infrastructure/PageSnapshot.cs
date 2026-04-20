using System;
using System.Collections.Generic;

namespace SPUtil.Infrastructure
{
    // ═══════════════════════════════════════════════════════════════════════════
    //  PageSnapshot — полный снимок Publishing-страницы
    //  Используется для чтения, клонирования и восстановления страниц.
    // ═══════════════════════════════════════════════════════════════════════════
    public class PageSnapshot
    {
        /// <summary>Server-relative URL страницы, e.g. /sites/hr/Pages/Home.aspx</summary>
        public string PageRelativeUrl { get; set; } = string.Empty;

        /// <summary>Title страницы (поле Title элемента списка)</summary>
        public string PageTitle { get; set; } = string.Empty;

        /// <summary>
        /// Server-relative URL макета страницы (PageLayout).
        /// e.g. /_catalogs/masterpage/ArticleLeft.aspx
        /// Нужен при создании новой страницы.
        /// </summary>
        public string LayoutRelativeUrl { get; set; } = string.Empty;

        /// <summary>
        /// Сырой HTML поля PublishingPageContent как он хранится в списке.
        /// Содержит ms-rte-wpbox заглушки с ZoneKey GUID.
        /// При клонировании ZoneKey заменяются на новые.
        /// </summary>
        public string PublishingHtml { get; set; } = string.Empty;

        /// <summary>
        /// WebParts в визуальном порядке (порядок ZoneKey в PublishingHtml).
        /// WebParts вне PublishingContent (в именованных зонах) — в конце списка.
        /// </summary>
        public List<WebPartSnapshot> WebParts { get; set; } = new();

        /// <summary>Когда был сделан снимок</summary>
        public DateTime SnapshotTime { get; set; } = DateTime.Now;
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  WebPartSnapshot — снимок одного WebPart
    // ═══════════════════════════════════════════════════════════════════════════
    public class WebPartSnapshot
    {
        /// <summary>
        /// StorageKey = WebPartDefinition.Id из LimitedWebPartManager.
        /// Атрибут webpartid="..." в рендеренном HTML страницы.
        /// Используется для операций через CSOM (удаление, экспорт XML).
        /// </summary>
        public string StorageKey { get; set; } = string.Empty;

        /// <summary>
        /// ZoneKey = webpartid2="..." в рендеренном HTML страницы.
        /// Используется в div_GUID заглушках в PublishingPageContent.
        /// Это НЕ то же самое что StorageKey.
        /// </summary>
        public string ZoneKey { get; set; } = string.Empty;

        /// <summary>
        /// Позиция в PublishingContent (1-based).
        /// 0 = WebPart находится в именованной зоне вне PublishingContent.
        /// </summary>
        public int VisualPosition { get; set; }

        /// <summary>ZoneId в LimitedWebPartManager (обычно "wpz" для Publishing)</summary>
        public string ZoneId { get; set; } = string.Empty;

        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// Полный .webpart XML полученный через exportwp.aspx.
        /// Содержит ВСЕ свойства включая кастомные.
        /// Используется для ImportWebPart при клонировании.
        /// </summary>
        public string ExportXml { get; set; } = string.Empty;

        /// <summary>
        /// Свойства из LimitedWebPartManager.WebPart.Properties.
        /// Используется для отображения и точечного обновления свойств.
        /// </summary>
        public Dictionary<string, string> Properties { get; set; } = new();
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  WebPartUpdateRequest — запрос на изменение WebPart
    //  Используется в методах UpdateWebPart, UpdateAllWebParts, UpdateProperty.
    // ═══════════════════════════════════════════════════════════════════════════
    public class WebPartUpdateRequest
    {
        /// <summary>StorageKey WebPart который нужно изменить</summary>
        public string StorageKey { get; set; } = string.Empty;

        /// <summary>Новый Title (null = не менять)</summary>
        public string? NewTitle { get; set; }

        /// <summary>Свойства которые нужно обновить. Key = имя свойства, Value = новое значение.</summary>
        public Dictionary<string, string> PropertiesToUpdate { get; set; } = new();

        /// <summary>
        /// Новая визуальная позиция в PublishingContent (1-based).
        /// null = не менять порядок.
        /// </summary>
        public int? NewVisualPosition { get; set; }
    }
}
