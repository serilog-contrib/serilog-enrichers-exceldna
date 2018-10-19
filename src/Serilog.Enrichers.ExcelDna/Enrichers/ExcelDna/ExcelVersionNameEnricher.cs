// Copyright 2018 Caio Proiete & Contributors
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
using ExcelDna.Integration;
using Serilog.Core;
using Serilog.Events;

namespace Serilog.Enrichers.ExcelDna
{
    /// <summary>
    /// Enriches log events with an ExcelVersionName property containing a string with the Excel Version
    /// based on the value from <see cref="ExcelDnaUtil.ExcelVersion"/>. E.g. "Excel 2016".
    /// </summary>
    public class ExcelVersionNameEnricher : ILogEventEnricher
    {
        private readonly bool _includeExcelBitness;
        private LogEventProperty _cachedProperty;

        /// <summary>
        /// Cretes an instance of the <see cref="ExcelVersionNameEnricher"/>.
        /// </summary>
        /// <param name="includeExcelBitness">Include (or not) the bitness of Excel in the name. E.g. Excel 2016 32-bit</param>
        public ExcelVersionNameEnricher(bool includeExcelBitness = true)
        {
            _includeExcelBitness = includeExcelBitness;
        }

        /// <summary>
        /// The property name added to enriched log events.
        /// </summary>
        public const string ExcelVersionNamePropertyName = "ExcelVersionName";

        /// <inheritdoc />
        public void Enrich(LogEvent logEvent, ILogEventPropertyFactory propertyFactory)
        {
            var cachedProperty =
                _cachedProperty =
                    _cachedProperty ?? propertyFactory.CreateProperty(ExcelVersionNamePropertyName, GetExcelVersionName(_includeExcelBitness));

            logEvent.AddPropertyIfAbsent(cachedProperty);
        }

        private static string GetExcelVersionName(bool includeExcelBitness)
        {
            var versionNumber = Convert.ToInt32(ExcelDnaUtil.ExcelVersion);
            string versionName;

            switch (versionNumber)
            {
                case 17: // Office 2019 - 17.0
                    versionName = "Excel 2019";
                    break;

                case 16: // Office 2016 - 16.0
                    versionName = "Excel 2016";
                    break;

                case 15: // Office 2013 - 15.0
                    versionName = "Excel 2013";
                    break;

                case 14: // Office 2010 - 14.0
                    versionName = "Excel 2010";
                    break;

                case 12: // Office 2007 - 12.0
                    versionName = "Excel 2007";
                    break;

                case 11: // Office 2003 - 11.0
                    versionName = "Excel 2003";
                    break;

                default:
                {
                    versionName = versionNumber < 11 ? "Excel < 2003" : "Excel > 2019";
                    break;
                }
            }

            return includeExcelBitness ? $"{versionName} {GetExcelBitnessInfo()}" : versionName;
        }

        private static string GetExcelBitnessInfo()
        {
            return Environment.Is64BitProcess ? "64-bit" : "32-bit";
        }
    }
}
