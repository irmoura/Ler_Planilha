/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package lendoxlsx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author irmoura
 */
public class LendoXLSX {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        
        FileInputStream fisPlanilha = null;
        
        try {
                
            File file = new File("C:\\planilhas\\planilhaDaAula.xlsx");
            fisPlanilha = new FileInputStream(file);
            
            /*CRIA UM WORKBOOK = PLANILHA TODA COM TODAS AS ABAS*/
            XSSFWorkbook workbook = new XSSFWorkbook(fisPlanilha);
            
            /*RECUPERAMOS APENAS A PRIMEIRA ABA OU PRIMEIRA PLANILHA*/
            XSSFSheet sheet = workbook.getSheetAt(0);
            
            /*RETORNA TODAS AS LINHAS DA PLANILHA 0 */
            Iterator<Row> rowIterator = sheet.iterator();
            
            /*VARRE TODAS AS LINHAS DA PLANILHA 0*/
            while(rowIterator.hasNext()){
                
                //recebe cada linha da planilha
                Row row = rowIterator.next();
                
                //pegamos todas as celulas desta linha
                Iterator<Cell> cellIterator = row.iterator();
                
                //varremos todas as celulas da linha atual
                while(cellIterator.hasNext()){
                    
                    //criamos uma celula
                    Cell cell = cellIterator.next();
                    
                    switch(cell.getCellType()){
                        
                        case Cell.CELL_TYPE_STRING:
                            System.out.println("TIPO STRING: "+cell.getStringCellValue());
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.println("TIPO NUMERICO: "+cell.getNumericCellValue());
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            System.out.println("TIPO FORMULA: "+cell.getCellFormula());
                            break;
                        
                    }
                    
                }
                
            }
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(LendoXLSX.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(LendoXLSX.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                fisPlanilha.close();
            } catch (IOException ex) {
                Logger.getLogger(LendoXLSX.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        
    }
    
}
