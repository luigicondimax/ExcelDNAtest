using Avalonia.Threading;
using ExcelDna.Integration.CustomUI;
using ExcelDNAtest.Resources;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using System.Diagnostics;
using System.Resources;
using System.Runtime.InteropServices;


namespace ExcelDNAtest
{
    // https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa722523(v=office.12)

    [ComVisible(true)]
    public class RibbonController :
#if AOT
    IExcelRibbon
#else
    ExcelRibbon
#endif    
    {
        internal static RibbonUITarget? myRibbon;

#if !AOT
        public override string GetCustomUI(string RibbonID)
#else
        public string GetCustomUI(string RibbonID)
#endif
        {

            // https://groups.google.com/g/exceldna/c/niFYofn2NYE
            try
            {

                return $"""
            <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='ribbonLoaded' loadImage='LoadImage'>
                <ribbon>
                    <tabs>
                        <tab id='tab2' label='Exceldna'>
                
                            <group id='Impostazioni' label='TEST'>
                                <button id='MostraImpostazioniExcelToolset' size='large' label='MessageBox' onAction='CreateMessageBox' imageMso='CurrentViewSettings'/>
                            </group>
                            <group id='Form' label='Form'>
                                <button id='OpenForm' size='large' label='{RibbonResources.Apri_la_form}' onAction='OpenForm' image="ExcelDNAtest.Image.Icon.ico" />
                            </group>
                            <group id='InvForm' label='InvalidateForm'>
                                <button id='InvalidateForm' label='InvalidateForm' onAction='InvalidateForm' image="ExcelDNAtest.Image.Package.ico"/>
                            </group>
                            <group id='Form4' label='Label'>
                                <button id='GetLabelId' size='large' getLabel='GetLabel' onAction='OpenForm' image="ExcelDNAtest.Image.Icon.ico" />
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
        public string GetLabel(RibbonControlTarget control)
        {
            return RibbonResources.Apri_la_form;
        }

        public void ribbonLoaded(RibbonUITarget ribbon)
        {
            // the only cast possible but is empty
            //if (ribbon is ExcelDna.Integration.CustomUI.RibbonControl rc)
            //{

            //}
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

                    Trace.TraceError($"Error in CreateMessageBox: {e.Message}");
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
                    var form = new FormSettings();
                    form.Show();
                }
                catch (Exception e)
                {

                    Trace.TraceError($"Error opening FormImpostazioniExcelToolset: {e.Message}");
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