using System;
using System.Collections.Generic;
using System.Text;

namespace SPUtil.Infrastructure
{
    public class SPListInfo
    {
        public string Title { get; set; } = string.Empty;
        public string InternalName { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string URL { get; set; } = string.Empty;
        public string ParentWebUrl { get; set; } = string.Empty;
        public string ServerRelativeUrl { get; set; } = string.Empty;
        public int ItemCount { get; set; }
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        public int BaseTemplate { get; set; } // Важно для логики копирования (100 или 101)

        // Переопределим ToString, чтобы старый код (где выводилась строка) не сломался сразу
        public override string ToString()
        {
            return $"URL:{URL}\n" +
                $"Display Name: {Title}\n" +
                   $"Internal Name: {InternalName}\n" +
                   $"Type: {Type}\n" +
                   $"BaseTemplate: {BaseTemplate.ToString()}\n" +
                   $"ParentWebUrl: {ParentWebUrl}\n" +
                   $"ServerRelativeUrl: {ServerRelativeUrl}\n" +
                   $"Items: {ItemCount}\n" +
                   $"Created: {Created.ToShortDateString()}\n"+
                   $"Modified: {Modified.ToShortDateString()}";
        }
    }
}
