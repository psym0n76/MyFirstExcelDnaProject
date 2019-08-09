using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace ExcelDnaLibrary
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
              <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
              <ribbon>
                <tabs>
                  <tab id='tab1' label='Custom Tab'>
                    <group id='group1' label='Functions'>
                      <button id='button1' label='GetData' onAction='OnButtonPressed'/>
                        <button id='button2' label='Filter Top 10' onAction='OnTopTenPressed'/>
                    </group >
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("Hello from control " + control.Id);
            MyExcelFunctions.WriteDataFromArrayToRange();
        }

        public void OnTopTenPressed(IRibbonControl control)
        {
            MyExcelFunctions.GetTopTen();
        }


    }
}