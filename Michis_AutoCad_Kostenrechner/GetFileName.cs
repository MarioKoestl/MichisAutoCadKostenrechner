// See https://aka.ms/new-console-template for more information
using System.Windows.Forms;

public class GetFileName
{
    public FileNameReturnObject Value()
    {
        string fileName = "";
        string safeFileName = "";
        OpenFileDialog dialog = new OpenFileDialog();
        dialog.InitialDirectory = @"C:\Users\Michal\AppData\Roaming\Autodesk\AutoCAD 2023\R24.2\enu\Recent\Save Drawing As";
        if (DialogResult.OK == dialog.ShowDialog())
        {
            fileName = dialog.FileName;
            safeFileName = dialog.SafeFileName;
        }
        return new FileNameReturnObject
        {
            FullFilePath = fileName,
            FileName = safeFileName
        };
    }
}
