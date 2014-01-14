using System.Data;
using System.Threading.Tasks;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;

public class Startup
{

    public async Task<object> Invoke(object input) {

        HSSFWorkbook hssfworkbook = new HSSFWorkbook();

        ////create a entry of DocumentSummaryInformation
        DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
        dsi.Company = "NPOI Team";
        hssfworkbook.DocumentSummaryInformation = dsi;

        ////create a entry of SummaryInformation
        SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
        si.Subject = "NPOI SDK Example";
        hssfworkbook.SummaryInformation = si;

        //here, we must insert at least one sheet to the workbook. otherwise, Excel will say 'data lost in file'
        //So we insert three sheet just like what Excel does
        hssfworkbook.CreateSheet("Sheet1");
        hssfworkbook.CreateSheet("Sheet2");
        hssfworkbook.CreateSheet("Sheet3");
        hssfworkbook.CreateSheet("Sheet4");

        ((HSSFSheet)hssfworkbook.GetSheetAt(0)).AlternativeFormula = false;
        ((HSSFSheet)hssfworkbook.GetSheetAt(0)).AlternativeExpression = false;

        //Write the stream data of workbook to the root directory
        FileStream file = new FileStream(@"test.xls", FileMode.Create);
        hssfworkbook.Write(file);
        file.Close();

        return null;
    }

    public async Task<object> OpenWorkbook(object input) {

        HSSFWorkbook hssfworkbook = new HSSFWorkbook();

        ////create a entry of DocumentSummaryInformation
        DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
        dsi.Company = "NPOI Team";
        hssfworkbook.DocumentSummaryInformation = dsi;

        ////create a entry of SummaryInformation
        SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
        si.Subject = "NPOI SDK Example";
        hssfworkbook.SummaryInformation = si;

        //here, we must insert at least one sheet to the workbook. otherwise, Excel will say 'data lost in file'
        //So we insert three sheet just like what Excel does
        hssfworkbook.CreateSheet("Sheet1");
        hssfworkbook.CreateSheet("Sheet2");
        hssfworkbook.CreateSheet("Sheet3");
        hssfworkbook.CreateSheet("Sheet4");

        ((HSSFSheet)hssfworkbook.GetSheetAt(0)).AlternativeFormula = false;
        ((HSSFSheet)hssfworkbook.GetSheetAt(0)).AlternativeExpression = false;

        //Write the stream data of workbook to the root directory
        FileStream file = new FileStream(@"test2.xls", FileMode.Create);
        hssfworkbook.Write(file);
        file.Close();

        return null;
    }

}