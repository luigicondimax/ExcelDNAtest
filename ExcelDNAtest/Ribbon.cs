using Avalonia.Threading;
using ExcelDna.Integration.CustomUI;
using ExcelDNAtest.Resources;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcelDNAtest
{
    // https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa722523(v=office.12)

#if AOT
    public class RibbonController :IExcelRibbon
#else
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
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
                // icon must be created from BMP and not from PNG. PNG --> BMP --> ICO
                return $"""
            <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='ribbonLoaded' loadImage='LoadImage'>
                <ribbon>
                    <tabs>
                        <tab id='tab2' label='Exceldna'>
                
                            <group id='Msg' label='Msg'>
                                <button id='MostraImpostazioniExcelToolset' size='large' label='MessageBox' onAction='CreateMessageBox' imageMso='CurrentViewSettings'/>
                            </group>
                            <group id='Setting' label='Setting'>
                                <button id='OpenSettingForm' size='large' label='{RibbonResources.Apri_la_form}' onAction='OpenSettingForm' image="ExcelDNAtest.Image.Icon.ico" />
                            </group>
                            <group id='InvForm' label='InvalidateForm'>
                                <button id='InvalidateForm' label='InvalidateForm' onAction='InvalidateForm' image="ExcelDNAtest.Image.Package.ico"/>
                            </group>
                            <group id='Form4' label=' Dynamic label'>
                                <labelControl id='GetLabelId' getLabel='GetLabel'/>
                            </group>
                            <group id='group1' getVisible='isFunzioniVisible' label='Test visibility'>
                                <labelControl id='dasd' label='Text hidden'/>
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
        public bool isFunzioniVisible(RibbonControlTarget control)
        {
            return true;
        }

        public void ribbonLoaded(RibbonUITarget ribbon)
        {
            // the only cast possible but is empty
            myRibbon = ribbon;
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

        private static FormSettings? _istanzaFormSettings;
        public async void OpenSettingForm(RibbonControlTarget control)
        {
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                try
                {
                    if (_istanzaFormSettings != null)
                    {
                        _istanzaFormSettings.Activate();
                        return;
                    }

                    _istanzaFormSettings = new FormSettings();

                    _istanzaFormSettings.Closed += (sender, e) =>
                    {
                        _istanzaFormSettings = null;
                    };

                    _istanzaFormSettings.Show();
                }
                catch (Exception e)
                {
                    Trace.TraceError($"Error opening FormSettings: {e.Message}");
                    _istanzaFormSettings = null;
                    throw;
                }
            });
        }
        public async void InvalidateForm(RibbonControlTarget control)
        {
            myRibbon?.Invalidate();
        }

    }
}
