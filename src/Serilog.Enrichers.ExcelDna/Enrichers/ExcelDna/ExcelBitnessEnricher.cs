// Copyright 2018-2019 Caio Proiete & Contributors
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

using System;
using Serilog.Core;
using Serilog.Events;

namespace Serilog.Enrichers.ExcelDna
{
    /// <summary>
    /// Enriches log events with an ExcelBitness property containing a string indicating if the Excel
    /// running your Excel-DNA add-in is 32-bit or 64-bit. E.g. "32-bit"
    /// </summary>
    public class ExcelBitnessEnricher : ILogEventEnricher
    {
        private LogEventProperty _cachedProperty;

        /// <summary>
        /// The property name added to enriched log events.
        /// </summary>
        public const string ExcelBitnessPropertyName = "ExcelBitness";

        /// <inheritdoc />
        public void Enrich(LogEvent logEvent, ILogEventPropertyFactory propertyFactory)
        {
            var cachedProperty =
                _cachedProperty =
                    _cachedProperty ?? propertyFactory.CreateProperty(ExcelBitnessPropertyName, GetExcelBitnessInfo());

            logEvent.AddPropertyIfAbsent(cachedProperty);
        }

        private static string GetExcelBitnessInfo()
        {
            return Environment.Is64BitProcess ? "64-bit" : "32-bit";
        }
    }
}
