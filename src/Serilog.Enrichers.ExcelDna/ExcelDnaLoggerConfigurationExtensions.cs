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
using ExcelDna.Integration;
using Serilog.Configuration;
using Serilog.Enrichers.ExcelDna;

namespace Serilog
{
    /// <summary>
    /// Extends <see cref="LoggerConfiguration"/> to add enrichers with properties from
    /// Excel-DNA.
    /// </summary>
    public static class ExcelDnaLoggerConfigurationExtensions
    {
        /// <summary>
        /// Enrich log events with an XllPath property containing the <see cref="ExcelDnaUtil.XllPath"/>.
        /// </summary>
        /// <param name="enrichmentConfiguration">Logger enrichment configuration.</param>
        /// <returns>Configuration object allowing method chaining.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static LoggerConfiguration WithXllPath(
            this LoggerEnrichmentConfiguration enrichmentConfiguration)
        {
            if (enrichmentConfiguration == null) throw new ArgumentNullException(nameof(enrichmentConfiguration));
            return enrichmentConfiguration.With(new XllPathEnricher());
        }

        /// <summary>
        /// Enrich log events with an ExcelVersion property containing the <see cref="ExcelDnaUtil.ExcelVersion"/>.
        /// </summary>
        /// <param name="enrichmentConfiguration">Logger enrichment configuration.</param>
        /// <returns>Configuration object allowing method chaining.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static LoggerConfiguration WithExcelVersion(
            this LoggerEnrichmentConfiguration enrichmentConfiguration)
        {
            if (enrichmentConfiguration == null) throw new ArgumentNullException(nameof(enrichmentConfiguration));
            return enrichmentConfiguration.With(new ExcelVersionEnricher());
        }

        /// <summary>
        /// Enrich log events with an ExcelVersionName property containing a string with the Excel Version
        /// based on the value from <see cref="ExcelDnaUtil.ExcelVersion"/>. E.g. "Excel 2016".
        /// </summary>
        /// <param name="enrichmentConfiguration">Logger enrichment configuration.</param>
        /// <param name="includeExcelBitness">Include (or not) the bitness of Excel in the name. E.g. Excel 2016 32-bit</param>
        /// <returns>Configuration object allowing method chaining.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static LoggerConfiguration WithExcelVersionName(
            this LoggerEnrichmentConfiguration enrichmentConfiguration, bool includeExcelBitness = true)
        {
            if (enrichmentConfiguration == null) throw new ArgumentNullException(nameof(enrichmentConfiguration));
            return enrichmentConfiguration.With(new ExcelVersionNameEnricher(includeExcelBitness));
        }

        /// <summary>
        /// Enrich log events with an ExcelBitness property containing a string indicating if the Excel
        /// running your Excel-DNA add-in is 32-bit or 64-bit. E.g. "32-bit"
        /// </summary>
        /// <param name="enrichmentConfiguration">Logger enrichment configuration.</param>
        /// <returns>Configuration object allowing method chaining.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static LoggerConfiguration WithExcelBitness(
            this LoggerEnrichmentConfiguration enrichmentConfiguration)
        {
            if (enrichmentConfiguration == null) throw new ArgumentNullException(nameof(enrichmentConfiguration));
            return enrichmentConfiguration.With(new ExcelBitnessEnricher());
        }
    }
}
