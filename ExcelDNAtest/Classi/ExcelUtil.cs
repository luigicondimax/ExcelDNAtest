#if RELEASEAOT
global using ExcelAppTarget = ExcelDna.Integration.IDynamic;
global using ExcelWorkbookTarget = ExcelDna.Integration.IDynamic;
global using ExcelWorksheetTarget = ExcelDna.Integration.IDynamic;
global using ExcelRangeTarget = ExcelDna.Integration.IDynamic;
global using ExcelChartTarget = ExcelDna.Integration.IDynamic;
global using RibbonControlTarget = ExcelDna.Integration.CustomUI.RibbonControl;
global using RibbonUITarget = System.Object;
#else
global using ExcelAppTarget = Microsoft.Office.Interop.Excel.Application;
global using ExcelWorkbookTarget = Microsoft.Office.Interop.Excel.Workbook;
global using ExcelWorksheetTarget = Microsoft.Office.Interop.Excel.Worksheet;
global using ExcelRangeTarget = Microsoft.Office.Interop.Excel.Range;
global using ExcelChartTarget = Microsoft.Office.Interop.Excel.Chart;
global using RibbonControlTarget = Microsoft.Office.Core.IRibbonControl;
global using RibbonUITarget = ExcelDna.Integration.CustomUI.IRibbonUI;

#endif

using ExcelDna.Integration;
using static ExcelDNAtest.Classi.Enum_Dict;


namespace ExcelDNAtest
{
    internal static class ExcelUtil
    {
        internal static ExcelReference Caller()
        {
            return XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
        }

        internal static string GetSheetName(ExcelReference reference)
        {
            if (reference == null) return string.Empty;
            string fullPath = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
            int index = fullPath.LastIndexOf(']');
            return index >= 0 ? fullPath.Substring(index + 1) : fullPath;
        }

        internal static ExcelRangeTarget ToRange(ExcelReference reference)
        {
            if (reference == null) return null!;

            ExcelAppTarget xlApp = ActiveApplication();
            string sheetName = GetSheetName(reference);

#if RELEASEAOT
            var sheets = (IDynamic)xlApp.Get("Sheets");
            var ws = (IDynamic)sheets.Get("Item", new object[] { sheetName });

            var cellaInizio = (IDynamic)ws.Get("Cells", new object[] { reference.RowFirst + 1, reference.ColumnFirst + 1 });
            var cellaFine = (IDynamic)ws.Get("Cells", new object[] { reference.RowLast + 1, reference.ColumnLast + 1 });

            return (ExcelRangeTarget)xlApp.Get("Range", new object[] { cellaInizio, cellaFine });
#else
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)xlApp.Sheets[sheetName];
            Microsoft.Office.Interop.Excel.Range cellaInizio = (Microsoft.Office.Interop.Excel.Range)ws.Cells[reference.RowFirst + 1, reference.ColumnFirst + 1];
            Microsoft.Office.Interop.Excel.Range cellaFine = (Microsoft.Office.Interop.Excel.Range)ws.Cells[reference.RowLast + 1, reference.ColumnLast + 1];

            return xlApp.get_Range(cellaInizio, cellaFine);
#endif
        }

        // =================================================================
        // METODI DI ACCESSO AGLI OGGETTI (INTEROP / AOT COMPATIBILI)
        // =================================================================

        internal static ExcelAppTarget ActiveApplication()
        {
#if RELEASEAOT
            return (ExcelAppTarget)ExcelDnaUtil.DynamicApplication;
#else
            return (ExcelAppTarget)ExcelDnaUtil.Application;
#endif

        }

        internal static ExcelWorksheetTarget ActiveSheet()
        {
#if RELEASEAOT
            return (ExcelWorksheetTarget)ActiveApplication().Get("ActiveSheet");
#else
            return ActiveApplication().ActiveSheet;
#endif
        }

        internal static ExcelWorkbookTarget ActiveWorkbook()
        {
#if RELEASEAOT
            return (ExcelWorkbookTarget)ActiveApplication().Get("ActiveWorkbook");
#else
            return ActiveApplication().ActiveWorkbook;
#endif
        }

        internal static ExcelRangeTarget ActiveCell()
        {
#if RELEASEAOT
            return (ExcelRangeTarget)ActiveApplication().Get("ActiveCell");
#else
            return ActiveApplication().ActiveCell;
#endif
        }

        internal static ExcelRangeTarget Range(string indirizzo)
        {
#if RELEASEAOT
            return (ExcelRangeTarget)ActiveApplication().Get("Range", new object[] { indirizzo });
#else
            return ActiveApplication().Range[indirizzo];
#endif
        }

        internal static ExcelChartTarget ActiveChart()
        {
#if RELEASEAOT
            return (ExcelChartTarget)ActiveApplication().Get("ActiveChart");
#else
            return ActiveApplication().ActiveChart;
#endif
        }

        internal static object Evaluate(string espressione)
        {
#if RELEASEAOT
            return ActiveApplication().Invoke("Evaluate", new object[] { espressione });
#else
            return ActiveApplication().Evaluate(espressione);
#endif
        }

        internal static bool TryGetSelectionAsRange(out ExcelRangeTarget selection)
        {
            var app = ActiveApplication();
#if RELEASEAOT
            var sel = app.Get("Selection");

            if (sel != null && (string)app.Invoke("TypeName", new object[] { sel }) == "Range")
            {
                selection = (ExcelRangeTarget)sel;
                return true;
            }
#else
            if (app.Selection is Microsoft.Office.Interop.Excel.Range sel)
            {
                selection = sel;
                return true;
            }
#endif
            selection = null!;
            return false;
        }

        internal static void SetScreenUpdating(bool enable)
        {
#if RELEASEAOT
            // Per il metodo .Set() di proprietà dirette, di solito basta il valore, 
            // ma se la firma richiede indicizzazione vuota si usa la sintassi standard:
            ActiveApplication().Set("ScreenUpdating", enable);
#else
            ActiveApplication().ScreenUpdating = enable;
#endif
        }

    }

#if RELEASEAOT
    internal enum XlChartType { xlXYScatterSmoothNoMarkers = 73 }
    internal enum XlSheetVisibility { xlSheetVisible = -1 }
    internal enum XlFixedFormatType { xlTypePDF = 0 }
#else
    internal enum XlChartType { xlXYScatterSmoothNoMarkers = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterSmoothNoMarkers }
    internal enum XlSheetVisibility { xlSheetVisible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible }
    internal enum XlFixedFormatType { xlTypePDF = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF }
#endif

    // =================================================================
    // METODI DI ESTENSIONE
    // =================================================================
    internal static class ExcelExtensions
    {
        internal static object GetValue(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            return range.Get("Value");
#else
            return range.Value;
#endif
        }

        internal static void SetValue(this ExcelRangeTarget range, object value)
        {
#if RELEASEAOT
            range.Set("Value", value);
#else
            range.Value = value;
#endif
        }

        internal static string GetText(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            return (string)range.Get("Text");
#else
            return range.Text;
#endif
        }

        internal static void SetNumberFormat(this ExcelRangeTarget range, string format)
        {
#if RELEASEAOT
            range.Set("NumberFormat", format);
#else
            range.NumberFormat = format;
#endif
        }

        // Lettura della Formula della cella
        internal static string GetFormula(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            return range.Get("Formula")?.ToString() ?? string.Empty;
#else
        return range.Formula?.ToString() ?? string.Empty;
#endif
        }

        // Scrittura della Formula della cella
        internal static void SetFormula(this ExcelRangeTarget range, object formula)
        {
#if RELEASEAOT
            range.Set("Formula", formula);
#else
            range.Formula = formula;
#endif
        }
        internal static string GetAddress(this ExcelRangeTarget range, bool external = false)
        {
#if RELEASEAOT
            // I parametri posizionali di Excel per Address sono:
            // 1. RowAbsolute (true)
            // 2. ColumnAbsolute (true)
            // 3. ReferenceStyle (1 = xlA1)
            // 4. External (il nostro valore bool)
            return (string)range.Get("Address", new object[] { true, true, 1, external });
#else
            return range.Address[External: external];
#endif
        }

        internal static object InputBox(string prompt, string title, object? defaultValue = null,
            XlInputBoxDataType type = XlInputBoxDataType.Text)
        {
            // Risolviamo il valore di default a tempo di esecuzione
            object finalDefault = defaultValue ?? Type.Missing;

#if RELEASEAOT
            return ExcelUtil.ActiveApplication().Invoke("InputBox", new object[]
            {
                prompt,
                title,
                finalDefault,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                (int)type // Castiamo l'enum all'intero sottostante per IDynamic
            });
#else
            // Per Interop dobbiamo convertire il nostro enum nell'enum nativo di Excel
            return ExcelUtil.ActiveApplication().InputBox(prompt, title, Default: finalDefault, Type: (int)type);
#endif
        }

        internal static double? OttieniNumeroDaInputBox(string prompt, string title)
        {
            // Sfruttiamo i parametri nominali di C# per saltare il defaultValue che è null
            object risultato = InputBox(prompt, title, type: XlInputBoxDataType.Number);

            // Se l'utente preme Annulla, Excel restituisce il bool false
            if (risultato is bool b && b == false) return null;

            // Essendo un InputBox numerico (Type 1), Excel restituisce già un double o un int boxed.
            // Intercettandolo direttamente evitiamo bug di formattazione stringa (punti/virgole)
            if (risultato is double d) return d;
            if (risultato is int i) return i;

            // Fallback di sicurezza estrema nel caso arrivasse come stringa sporca
            if (risultato != null && double.TryParse(risultato.ToString(), out double val)) return val;

            return null;
        }

        /// <summary>
        /// Restituisce un Range di Excel, o null se l'utente preme Annulla
        /// </summary>
        internal static ExcelRangeTarget? OttieniRangeDaInputBox(string prompt, string title, string indirizzoDefault = "")
        {
            object risultato = InputBox(prompt, title, indirizzoDefault, XlInputBoxDataType.Range);

            if (risultato is bool b && b == false) return null;
            return risultato as ExcelRangeTarget;
        }

        internal static ExcelRangeTarget GetOffset(this ExcelRangeTarget range, int rowOffset, int columnOffset)
        {
#if RELEASEAOT
            return (ExcelRangeTarget)range.Get("Offset", new object[] { rowOffset, columnOffset });
#else
        return range.Offset[rowOffset, columnOffset];
#endif
        }

        // Lettura e Scrittura dello Style
        internal static object GetStyle(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            return range.Get("Style");
#else
        return range.Style;
#endif
        }

        internal static void SetStyle(this ExcelRangeTarget range, object style)
        {
#if RELEASEAOT
            range.Set("Style", style);
#else
        range.Style = style;
#endif
        }

        // Lettura del NumberFormat (la scrittura l'avevamo già fatta!)
        internal static string GetNumberFormat(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            return (string)range.Get("NumberFormat");
#else
        return range.NumberFormat?.ToString() ?? "";
#endif
        }

        // Recupera il Foglio di lavoro (Parent) da un Range
        internal static ExcelWorksheetTarget GetParent(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            return (ExcelWorksheetTarget)range.Get("Parent");
#else
        return (ExcelWorksheetTarget)range.Parent;
#endif
        }

        // Nome del foglio di lavoro
        internal static string GetName(this ExcelWorksheetTarget sheet)
        {
#if RELEASEAOT
            return (string)sheet.Get("Name");
#else
        return sheet.Name;
#endif
        }

        // Attivazione del foglio
        internal static void Activate(this ExcelWorksheetTarget sheet)
        {
#if RELEASEAOT
            sheet.Invoke("Activate", Array.Empty<object>());
#else
            sheet.Activate();
#endif
        }

        // Selezione del Range
        internal static void Select(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            range.Invoke("Select", Array.Empty<object>());
#else
        range.Select();
#endif
        }

        // Creazione di un grafico sul foglio corrente
        internal static ExcelChartTarget AddChart(this ExcelWorksheetTarget sheet, XlChartType chartType)
        {
#if RELEASEAOT
            var shapes = (IDynamic)sheet.Get("Shapes");
            var chartShape = (IDynamic)shapes.Invoke("AddChart", new object[] { (int)chartType });
            return (ExcelChartTarget)chartShape.Get("Chart");
#else
        Microsoft.Office.Interop.Excel.Shape chartShape = sheet.Shapes.AddChart();
        Microsoft.Office.Interop.Excel.Chart chart = chartShape.Chart;
        chart.ChartType = (Microsoft.Office.Interop.Excel.XlChartType)chartType;
        return chart;
#endif
        }

        // Svuota in sicurezza tutte le serie esistenti di un grafico
        internal static void ClearSeries(this ExcelChartTarget chart)
        {
#if RELEASEAOT
            var seriesCollection = (IDynamic)chart.Invoke("SeriesCollection", Array.Empty<object>());
            int count = (int)seriesCollection.Get("Count");
            // Loop al contrario per evitare disallineamenti di indice durante l'eliminazione
            for (int i = count; i >= 1; i--)
            {
                var serie = (IDynamic)seriesCollection.Invoke("Item", new object[] { i });
                serie.Invoke("Delete", Array.Empty<object>());
            }
#else
        var seriesCollection = (Microsoft.Office.Interop.Excel.SeriesCollection)chart.SeriesCollection();
        foreach (Microsoft.Office.Interop.Excel.Series serie in seriesCollection)
        {
            serie.Delete();
        }
#endif
        }

        // Aggiunge una nuova serie al grafico usando formule stringa
        internal static void AddSeries(this ExcelChartTarget chart, string nameFormula, string xValuesFormula, string valuesFormula)
        {
#if RELEASEAOT
            var seriesCollection = (IDynamic)chart.Invoke("SeriesCollection", Array.Empty<object>());
            var nuovaSerie = (IDynamic)seriesCollection.Invoke("NewSeries", Array.Empty<object>());
            nuovaSerie.Set("Name", nameFormula);
            nuovaSerie.Set("XValues", xValuesFormula);
            nuovaSerie.Set("Values", valuesFormula);
#else
        var seriesCollection = (Microsoft.Office.Interop.Excel.SeriesCollection)chart.SeriesCollection();
        Microsoft.Office.Interop.Excel.Series nuovaSerie = seriesCollection.NewSeries();
        nuovaSerie.Name = nameFormula;
        nuovaSerie.XValues = xValuesFormula;
        nuovaSerie.Values = valuesFormula;
#endif
        }

#if RELEASEAOT
        // =================================================================
        // MODALITÀ NATIVE AOT
        // =================================================================
        // In AOT, sia le Application che i Worksheet sono 'IDynamic'.
        // Definiamo un UNICO metodo di estensione per evitare l'ambiguità di firma.
        internal static ExcelRangeTarget GetRange(this ExcelAppTarget target, string address)
        {
            return (ExcelRangeTarget)target.Get("Range", new object[] { address });
        }
#else
    // =================================================================
    // MODALITÀ INTEROP CLASSICA (NON-AOT)
    // =================================================================
    // In modalità classica i tipi sono ben distinti, quindi servono 
    // entrambi gli overload per consentire il forte legame del compilatore.
    internal static ExcelRangeTarget GetRange(this Microsoft.Office.Interop.Excel.Application app, string address)
    {
        return app.Range[address];
    }

    internal static ExcelRangeTarget GetRange(this Microsoft.Office.Interop.Excel.Worksheet sheet, string address)
    {
        return sheet.Range[address];
    }
#endif

        // Cancella i commenti dal Range
        internal static void ClearComments(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            range.Invoke("ClearComments", Array.Empty<object>());
#else
        range.ClearComments();
#endif
        }

        // Aggiunge un commento al Range
        internal static void AddComment(this ExcelRangeTarget range, string text)
        {
#if RELEASEAOT
            range.Invoke("AddComment", new object[] { text });
#else
        range.AddComment(text);
#endif
        }
        // Gestione DisplayAlerts della Application
        internal static void SetDisplayAlerts(this ExcelAppTarget app, bool value)
        {
#if RELEASEAOT
            app.Set("DisplayAlerts", value);
#else
        app.DisplayAlerts = value;
#endif
        }

        // Gestione Visibility della Application
        internal static void SetVisible(this ExcelAppTarget app, bool value)
        {
#if RELEASEAOT
            app.Set("Visible", value);
#else
        app.Visible = value;
#endif
        }

        // Aggiunge un nuovo Workbook e lo restituisce
        internal static ExcelWorkbookTarget AddWorkbook(this ExcelAppTarget app)
        {
#if RELEASEAOT
            var workbooks = (IDynamic)app.Get("Workbooks");
            return (ExcelWorkbookTarget)workbooks.Invoke("Add", Array.Empty<object>());
#else
        return app.Workbooks.Add();
#endif
        }

        // Recupera l'ActiveSheet corrente
        internal static ExcelWorksheetTarget GetActiveSheet(this ExcelAppTarget app)
        {
#if RELEASEAOT
            return (ExcelWorksheetTarget)app.Get("ActiveSheet");
#else
            return (ExcelWorksheetTarget)app.ActiveSheet;
#endif
        }

        // Recupera l'ActiveSheet corrente
        internal static ExcelRangeTarget GetActiveCell(this ExcelAppTarget app)
        {
#if RELEASEAOT
            return (ExcelRangeTarget)app.Get("ActiveCell");
#else
            return (ExcelRangeTarget)app.ActiveCell;
#endif
        }


        // Apre un Workbook gestendo i parametri posizionali in AOT
        internal static ExcelWorkbookTarget OpenWorkbook(this ExcelAppTarget app, string file, bool readOnly, bool ignoreReadOnlyRecommended)
        {
#if RELEASEAOT
            var workbooks = (IDynamic)app.Get("Workbooks");
            // Firma posizionale di Workbooks.Open: 1.Filename, 2.UpdateLinks, 3.ReadOnly, 4.Format, 5.Password, 6.WriteResPassword, 7.IgnoreReadOnlyRecommended
            return (ExcelWorkbookTarget)workbooks.Invoke("Open", new object[]
            {
            file,
            Type.Missing,
            readOnly,
            Type.Missing,
            Type.Missing,
            Type.Missing,
            ignoreReadOnlyRecommended
            });
#else
        return app.Workbooks.Open(file, ReadOnly: readOnly, IgnoreReadOnlyRecommended: ignoreReadOnlyRecommended);
#endif
        }

        // Recupera un foglio tramite Indice (int) o Nome (string)
        internal static ExcelWorksheetTarget GetSheet(this ExcelWorkbookTarget wb, object indexOrName)
        {
#if RELEASEAOT
            var sheets = (IDynamic)wb.Get("Sheets");
            return (ExcelWorksheetTarget)sheets.Get("Item", new object[] { indexOrName });
#else
        return (ExcelWorksheetTarget)wb.Sheets[indexOrName];
#endif
        }

        // Recupera la proprietà FullName del Workbook
        internal static string GetFullName(this ExcelWorkbookTarget wb)
        {
#if RELEASEAOT
            return (string)wb.Get("FullName");
#else
        return wb.FullName;
#endif
        }

        // Recupera una cella specifica per riga e colonna da un foglio
        internal static ExcelRangeTarget GetCell(this ExcelWorksheetTarget sheet, int row, int column)
        {
#if RELEASEAOT
            return (ExcelRangeTarget)sheet.Get("Cells", new object[] { row, column });
#else
        return (ExcelRangeTarget)sheet.Cells[row, column];
#endif
        }

        // Chiude il Workbook
        internal static void Close(this ExcelWorkbookTarget wb, bool saveChanges)
        {
#if RELEASEAOT
            wb.Invoke("Close", new object[] { saveChanges });
#else
        wb.Close(saveChanges);
#endif
        }

        // Imposta la visibilità del foglio di lavoro
        internal static void SetVisibility(this ExcelWorksheetTarget sheet, XlSheetVisibility visibility)
        {
#if RELEASEAOT
            sheet.Set("Visible", (int)visibility);
#else
        sheet.Visible = (Microsoft.Office.Interop.Excel.XlSheetVisibility)visibility;
#endif
        }

        // Recupera il ColorIndex del Font come stringa safely
        internal static string GetFontColorIndex(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            var font = (IDynamic)range.Get("Font");
            return font.Get("ColorIndex")?.ToString() ?? "";
#else
        return range.Font.ColorIndex?.ToString() ?? "";
#endif
        }

        // Recupera il ColorIndex dell'Interior (riempimento) come stringa safely
        internal static string GetInteriorColorIndex(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            var interior = (IDynamic)range.Get("Interior");
            return interior.Get("ColorIndex")?.ToString() ?? "";
#else
        return range.Interior.ColorIndex?.ToString() ?? "";
#endif
        }
        // Aggiunge un collegamento ipertestuale al Range
        internal static void AddHyperlink(this ExcelRangeTarget range, string fileAddress, string screenTip, string textToDisplay)
        {
#if RELEASEAOT
            var hyperlinks = (IDynamic)range.Get("Hyperlinks");
            // I parametri posizionali di Excel per Hyperlinks.Add sono:
            // 1. Anchor (range)
            // 2. Address (fileAddress)
            // 3. SubAddress (Type.Missing)
            // 4. ScreenTip (il tooltip al passaggio del mouse)
            // 5. TextToDisplay (il testo visibile nella cella)
            hyperlinks.Invoke("Add", new object[]
            {
            range,
            fileAddress,
            Type.Missing,
            screenTip,
            textToDisplay
            });
#else
        range.Hyperlinks.Add(range, fileAddress, Type.Missing, screenTip, textToDisplay);
#endif
        }
        internal static ExcelRangeTarget? GetSelection(this ExcelAppTarget app)
        {
#if RELEASEAOT
            object selection = app.Get("Selection");
            return selection as ExcelRangeTarget;
#else
        return app.Selection as Microsoft.Office.Interop.Excel.Range;
#endif
        }

        // Esegue un'azione su ogni singola cella del range evitando i problemi di enumerazione COM in AOT
        internal static void ForEachCell(this ExcelRangeTarget range, Action<ExcelRangeTarget> action)
        {
#if RELEASEAOT
            // Otteniamo il conteggio di righe e colonne per fare un ciclo for numerico pulito
            var rowsObj = (IDynamic)range.Get("Rows");
            int rows = (int)rowsObj.Get("Count");

            var colsObj = (IDynamic)range.Get("Columns");
            int cols = (int)colsObj.Get("Count");

            var cellsObj = (IDynamic)range.Get("Cells");

            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    var cell = (ExcelRangeTarget)cellsObj.Get("Item", new object[] { r, c });
                    action(cell);
                }
            }
#else
        foreach (Microsoft.Office.Interop.Excel.Range cell in range.Cells)
        {
            action(cell);
        }
#endif
        }

        internal static bool GetHasFormula(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            return (bool)(range.Get("HasFormula") ?? false);
#else
        return (bool)range.HasFormula;
#endif
        }

        // Recupera l'ActiveWorkbook corrente dell'applicazione
        internal static ExcelWorkbookTarget? GetActiveWorkbook(this ExcelAppTarget app)
        {
#if RELEASEAOT
            object wb = app.Get("ActiveWorkbook");
            return wb as ExcelWorkbookTarget;
#else
        return app.ActiveWorkbook;
#endif
        }

        // Esporta il Workbook mappando i parametri posizionali esatti di Excel
        internal static void ExportAsFixedFormat(this ExcelWorkbookTarget wb, XlFixedFormatType type, string filename, bool includeDocProperties, bool ignorePrintAreas, bool openAfterPublish)
        {
#if RELEASEAOT
            // Firma COM posizionale di ExportAsFixedFormat:
            // 1.Type, 2.Filename, 3.Quality, 4.IncludeDocProperties, 5.IgnorePrintAreas, 6.From, 7.To, 8.OpenAfterPublish
            wb.Invoke("ExportAsFixedFormat", new object[]
            {
            (int)type,
            filename,
            Type.Missing, // Quality
            includeDocProperties,
            ignorePrintAreas,
            Type.Missing, // From
            Type.Missing, // To
            openAfterPublish
            });
#else
        wb.ExportAsFixedFormat((Microsoft.Office.Interop.Excel.XlFixedFormatType)type, Filename: filename, IncludeDocProperties: includeDocProperties, IgnorePrintAreas: ignorePrintAreas, OpenAfterPublish: openAfterPublish);
#endif
        }

        // Invia in stampa il Workbook rispettando la posizione dei parametri COM
        internal static void PrintOut(this ExcelWorkbookTarget wb, int copies, bool collate)
        {
#if RELEASEAOT
            // Firma COM posizionale di PrintOut:
            // 1.From, 2.To, 3.Copies, 4.Preview, 5.ActivePrinter, 6.PrintToFile, 7.Collate
            wb.Invoke("PrintOut", new object[]
            {
            Type.Missing, // From
            Type.Missing, // To
            copies,
            Type.Missing, // Preview
            Type.Missing, // ActivePrinter
            Type.Missing, // PrintToFile
            collate
            });
#else
        wb.PrintOut(Copies: copies, Collate: collate);
#endif
        }

        // Forzza lo stato della cella/range come "da ricalcolare" (Dirty)
        internal static void Dirty(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            range.Invoke("Dirty", Array.Empty<object>());
#else
        range.Dirty();
#endif
        }

        // Esegue il ricalcolo immediato della cella/range
        internal static void Calculate(this ExcelRangeTarget range)
        {
#if RELEASEAOT
            range.Invoke("Calculate", Array.Empty<object>());
#else
        range.Calculate();
#endif
        }

        public static void InvalidateRibbon(this RibbonUITarget? ribbon)
        {
            if (ribbon == null) return;

#if RELEASEAOT
            if (ribbon is ExcelDna.Integration.IDynamic dynamicRibbon)
            {
                dynamicRibbon.Invoke("Invalidate", Array.Empty<object>());
            }
            else
            {
                System.Diagnostics.Trace.WriteLine("La Ribbon non implementa IDynamic.");
            }
            #else
        ribbon.Invalidate();
#endif
        }

        public static void InvalidateControlRibbon(this RibbonUITarget? ribbon, string controlId)
        {
            if (ribbon == null) return;

#if RELEASEAOT
            if (ribbon is ExcelDna.Integration.IDynamic dynamicRibbon)
            {
                dynamicRibbon.Invoke("InvalidateControl", new object[] { controlId });
            }
            else
            {
                System.Diagnostics.Trace.WriteLine("La Ribbon non implementa IDynamic.");
            }
#else
        ribbon.InvalidateControl(controlId);
#endif
        }
    }
}