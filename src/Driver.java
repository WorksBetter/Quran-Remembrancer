
import java.io.FileInputStream;  
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
import org.apache.poi.ss.usermodel.DataFormatter;



public class Driver  
{  
    public static void main(String[] args)   
    {  
        
        Driver rc = new Driver(); 
        XSSFWorkbook wb;
        
        try {
            wb = rc.getWorkbook("completed work.xlsx");

            if(wb != null){
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    Sheet sheet = wb.getSheetAt(i);
                   
                    System.out.println("Sheet Name: " + wb.getSheetName(i));
                    System.out.println("No.: " + Integer.toString(i));
                    
                    String vOutput = rc.ReadCellData(sheet ,0, 1);    
                    System.out.println(vOutput); 
               }
                
                Scanner sc = new Scanner(System.in);
                System.out.println("Enter Ayah number to view details: ");
                
                int ayahNumber = sc.nextInt();
                Sheet selectedAyahSheet;
                try {
                    selectedAyahSheet = wb.getSheetAt(ayahNumber);
                    rc.displayAyahInfo(selectedAyahSheet);
                } catch (IllegalArgumentException e) {
                    e.printStackTrace();
                }   
            }
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        
    }
    
    public XSSFWorkbook getWorkbook(String wbName) throws IOException{
        FileInputStream fis = null;
        XSSFWorkbook wb = null;
        try {
            //reading data from a file in the form of bytes
            fis = new FileInputStream(wbName); 
            
            //constructs an XSSFWorkbook object, by buffering the whole stream into the memory
            wb = new XSSFWorkbook(fis);
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
            return null;
        } finally {
            try {
                if(fis != null)
                    fis.close();
            } catch (IOException ex) {
                ex.printStackTrace();
                return null;
            }
        }
        return wb;
    }
    
    public String ReadCellData(Sheet sheet, int vRow, int vColumn)  
    {  
        DataFormatter formatter = new DataFormatter();
        

        Row row = sheet.getRow(vRow); //returns the logical row  
        Cell cell = row.getCell(vColumn); //getting the cell representing the given column
        
        String val = formatter.formatCellValue(cell);
        return val;               //returns the cell value  
    }
    
    
    public void displayAyahInfo(Sheet sheet){
        String urduTranslation;
        String reference;
        DataFormatter formatter = new DataFormatter();
        
        String ayat = ReadCellData(sheet, 0, 1);
        urduTranslation = formatter.formatCellValue(sheet.getRow(0).getCell(3));
        String numberOfReferences = ayat.replaceAll("[^0-9]", "");
        
        String[] referenceList = new String[Integer.valueOf(numberOfReferences + 1)];
       
        
        
        for (int i = 0; i < Integer.valueOf(numberOfReferences); i++) {  
            Row referenceRow = sheet.getRow(i + 1);
            Cell cell = referenceRow.getCell(1);
            
            
            if(cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK){
                reference = formatter.formatCellValue(cell);
                referenceList[i] = reference;
            }          
        }
        
        System.out.println("Urdu Translation: " + urduTranslation);
        
        System.out.println("List of References: ");
        for (String string : referenceList) {
            if(string != null)
                System.out.println(string);
        }
    }
}