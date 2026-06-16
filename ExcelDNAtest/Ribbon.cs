using ExcelDna.Integration.CustomUI;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Avalonia.Threading;


namespace ExcelDNAtest
{
    // https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa722523(v=office.12)

    [ComVisible(true)]
    public class RibbonController :
#if RELEASEAOT
    IExcelRibbon
#else
    ExcelRibbon
#endif    
    {
        internal static RibbonUITarget? myRibbon;

#if !RELEASEAOT
        public override string GetCustomUI(string RibbonID)
#else
        public string GetCustomUI(string RibbonID)
#endif
        {

            // Inserito all'interno di un blocco try-catch preventivo per i requisiti Excel-DNA https://groups.google.com/g/exceldna/c/niFYofn2NYE
            try
            {

                return $"""
            <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='ribbonLoaded' loadImage='CaricaImmaginiPersonalizzate' >
                <ribbon>
                    <tabs>
                        <tab id='tab2' label='Exceldna'>
                
                            <group id='Impostazioni' label='TEST'>
                                <button id='MostraImpostazioniExcelToolset' size='large' label='MessageBox' onAction='CreateMessageBox' imageMso='CurrentViewSettings'/>
                            </group>
                            <group id='Form' label='Form'>
                                <button id='OpenForm' size='large' label='OpenForm' onAction='OpenForm' imageMso='CurrentViewSettings'/>
                            </group>
                            <group id='InvForm' label='InvalidateForm'>
                                <button id='InvalidateForm' label='InvalidateForm' onAction='InvalidateForm'/>
                            </group>
                     
                        </tab>
                    </tabs>
                </ribbon>
            </customUI>
            """;
            }
            catch (Exception e)
            {
                Trace.TraceError($"Errore in GetCustomUI: {e.ToString()}");
                return "";
            }
        }

        public void ribbonLoaded(RibbonUITarget ribbon)
        {
            // the only cast possible but is empty
            if (ribbon is ExcelDna.Integration.CustomUI.RibbonControl rc)
            {

            }
                myRibbon =ribbon;
        }

        public async void CreateMessageBox(RibbonControlTarget control)
        {
            await Dispatcher.UIThread.InvokeAsync(async () =>
            {
                try
                {
                    var box = MessageBoxManager.GetMessageBoxStandard("Title", "Message", ButtonEnum.Ok, Icon.Warning);

                    await box.ShowWindowAsync();
                }
                catch (Exception e)
                {

                    Trace.TraceError($"Errore in CreateMessageBox: {e.Message}");
                    throw;
                }
            });
        }

        public async void OpenForm(RibbonControlTarget control)
        {
            await Dispatcher.UIThread.InvokeAsync(async () =>
            {
                try
                {
                    var form = new FormImpostazioniExcelToolset();
                    form.Show();
                }
                catch (Exception e)
                {

                    Trace.TraceError($"Errore in apertura FormImpostazioniExcelToolset: {e.Message}");
                    throw;
                }
            });
        }
        public async void InvalidateForm(RibbonControlTarget control)
        {
            myRibbon.InvalidateRibbon();
        }

    }
}