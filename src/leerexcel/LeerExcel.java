/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package leerexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.ss.formula.functions.Rows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Brandon Rogerio Aguirre Mendoza bran1189@gmail.com
 */
public class LeerExcel {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {

        //Crear archivo
        /*Workbook libro = new XSSFWorkbook();
        Sheet hoja = libro.createSheet("java");
        
        try {
            
            FileOutputStream archivo = new FileOutputStream(
                    new File("Reporte.xlsx"));
            libro.write(archivo);
            archivo.close();
            
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "No se pudo crear el archivo");
        }*/
        try {
            FileInputStream archivo = new FileInputStream(
                    "D:\\CursosUdemy\\CRUD en Java y MySQL facil\\LeerExcel\\Reporte.ods");
            XSSFWorkbook libro = new XSSFWorkbook(archivo);
            XSSFSheet hoja = libro.getSheetAt(0);

            int nFilas = hoja.getLastRowNum();
            for (int i = 0; i <= nFilas; i++) {
                Row fila = hoja.getRow(i);
                int nColumnas = fila.getLastCellNum();

                for (int j = 0; j < nColumnas; j++) {
                    Cell celda = fila.getCell(j);

                    switch (celda.getCellTypeEnum().toString()) {

                        case "NUMERIC":
                            System.out.print(celda.getNumericCellValue()
                                    + " ");
                            break;
                        case "STRING":
                            System.out.print(celda.getStringCellValue()
                                    + " ");
                            break;
                        case "FORMULA":
                            System.out.print(celda.getCellFormula()
                                    + " ");
                            break;

                    }

                }
            }

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Error al leer el archivo");
        }

    }
}
