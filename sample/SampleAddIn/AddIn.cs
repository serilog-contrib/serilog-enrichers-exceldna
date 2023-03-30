#region Copyright 2018-2023 C. Augusto Proiete & Contributors
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
//
#endregion

using System;
using ExcelDna.Integration;
using ExcelDna.Integration.Extensibility;
using ExcelDna.Logging;
using Serilog;

namespace SampleAddIn
{
    public class AddIn : ExcelComAddIn, IExcelAddIn
    {
        public void AutoOpen()
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Verbose()
                .Enrich.WithXllPath()
                .Enrich.WithExcelVersion()
                .Enrich.WithExcelVersionName()
                .Enrich.WithExcelBitness()
                .WriteTo.ExcelDnaLogDisplay(displayOrder: DisplayOrder.NewestFirst,
                    outputTemplate: "{Properties:j}{NewLine}[{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("Hello from {AddInName}! :)", DnaLibrary.CurrentLibrary.Name);
            LogDisplay.Show();

            ExcelComAddInHelper.LoadComAddIn(this);
        }

        public void AutoClose()
        {
            // Do nothing
        }

        public override void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            Log.Information("Goodbye from {AddInName}! :)", DnaLibrary.CurrentLibrary.Name);
            Log.CloseAndFlush();

            base.OnDisconnection(disconnectMode, ref custom);
        }
    }
}
