module MyRibbon

open System.Windows.Forms
open System.Runtime.InteropServices
open ExcelDna.Integration
open ExcelDna.Integration.CustomUI

// This defines a regular Excel macro (in Excel you can press Alt + F8, type in the name "showMessage", then click the Run button).
// For the ribbon, it will be run through the ExcelRibbon.RunTagMacro(...) helper, which run whatever macro is specified in the button tag attribute
// One advantage is that you can 
[<ExcelCommand>]
let showMessage () =
    XlCall.Excel(XlCall.xlcAlert, "Hello from a macro!") 
    |> ignore


// This type defines the ribbon interface. It is a public class that derives from ExcelRibbon
[<ComVisible(true)>]    // This attribute is only needed if there is an assembly-level [<assembly:ComVisible(false)>] attribute.
type public MyRibbon() =
    inherit ExcelRibbon()

    // The ribbon xml definition could also be placed in the .dna file
    // Remember to switch on the ExcelOption "Show add-in user interface errors" option (under the Advanced tab under General)
    override this.GetCustomUI(ribbonId) = 
        @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' >
          <ribbon>
            <tabs>
              <tab id='CustomTab' label='FSfull'>
                <group id='SampleGroup' label='My Sample Group'>
                  <button id='Button1' label='Run a macro' onAction='RunTagMacro' tag='showMessage' />
                  <button id='Button2' label='Run a class member' onAction='OnButtonPressed'/>
                </group >
              </tab>
            </tabs>
          </ribbon>
        </customUI>"

    member this.OnButtonPressed (control:IRibbonControl) =
        MessageBox.Show "Hello from F#!" 
        |> ignore

