using System.Collections.Generic;

namespace AutoOpener.Core.Models
{
    /// <summary>
    /// Описывает отдельную модель и её рабочие наборы для открытия
    /// </summary>
    public class ModelTask
    {
        /// <summary>
        /// Путь к модели (RSN://... или локальный/сетевой путь)
        /// </summary>
        public string ModelPath { get; set; }

        /// <summary>
        /// Список рабочих наборов для открытия
        /// </summary>
        public List<string> WorksetsByName { get; set; }

        public ModelTask()
        {
            WorksetsByName = new List<string>();
        }
    }
}
