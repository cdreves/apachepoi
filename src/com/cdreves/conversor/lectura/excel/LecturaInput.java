package com.cdreves.conversor.lectura.excel;

import com.cdreves.conversor.escritura.archivo.EscrituraArchivo;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.cdreves.conversor.procesosglobales.ProcesoGlobal;
import com.cdreves.conversor.utils.ExcelUtils;
import org.apache.commons.lang3.StringUtils;

/**
 * Clase para lectura de archivo Excel

 *
 * @author Carlos Dreves N
 * @version 1.0.0 28-07-2014.
 * @since JDK5.0
 */
public class LecturaInput {

    ProcesoGlobal procesosGlobales = new ProcesoGlobal();
    EscrituraArchivo escrituraArchivo = new EscrituraArchivo();

    /**
     * Metodo principal para lectura de archivo Excel con formato Xlsx.

     *
     * @param fileName Nombre de archivo excel a convertir.
     * @param numeroHoja N�mero de Hoja de archivo excel a convertir.
     * @see
     * powerdata.conversor.escritura.archivo.EscrituraArchivo#escrituraXlsxToCsv()
     *
     *
     */
    public void leerArchivoXlsx(String fileName, int numeroHoja,
            String OutputDir) {

        List cellDataList = new ArrayList();

        try {
            File file = new File(fileName);
            file.getName();
            OPCPackage opcPackage = OPCPackage.open(file);
            XSSFWorkbook workbook = new XSSFWorkbook(opcPackage);
            XSSFSheet xssfSheet = workbook.getSheetAt(numeroHoja);
            Iterator rowIterator = xssfSheet.rowIterator();

            while (rowIterator.hasNext()) {

                XSSFRow xssfRow = (XSSFRow) rowIterator.next();
                Iterator iterator = xssfRow.cellIterator();
                List cellTempList = new ArrayList();
                while (iterator.hasNext()) {
                    XSSFCell xssfCell = (XSSFCell) iterator.next();
                    String valor = ExcelUtils.getValue(xssfCell);

                    xssfCell.setCellType(Cell.CELL_TYPE_STRING);
                    System.out.println("Valor de celda: " + valor);

                    if (StringUtils.isEmpty(valor)) {
                        System.out.println("fila vacia");
                        //cellTempList.add(";");
                        cellTempList.add(xssfCell);
                    } else {
                        cellTempList.add(xssfCell);
                    }
                }
                cellDataList.add(cellTempList);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        escrituraArchivo.escrituraXlsxToCsv(cellDataList, OutputDir);





    }

    /**
     * Metodo principal para lectura de archivo Excel con formato Xls y que
     * ejecuta m�todo para escritura de archvo, enviandole por par�metro el
     * contenido del archivo.

     *
     * @param fileName Nombre de archivo excel a convertir.
     * @param numeroHoja N�mero de Hoja de archivo excel a convertir.
     * @see
     * powerdata.conversor.escritura.archivo.EscrituraArchivo#escrituraXlsToCsv()
     *
     *
     */
    public void leerArchivoXls(String fileName, int numeroHoja, String OutputDir) {

        List cellDataList = new ArrayList();

        try {
            FileInputStream fileInputStream = new FileInputStream(fileName);
            File file = new File(fileName);

            POIFSFileSystem fsFileSystem = new POIFSFileSystem(fileInputStream);
            HSSFWorkbook workBook = new HSSFWorkbook(fsFileSystem);
            HSSFSheet hssfSheet = workBook.getSheetAt(numeroHoja);
            Iterator rowIterator = hssfSheet.rowIterator();

            while (rowIterator.hasNext()) {
            
                //HSSFCell cell = null;
                HSSFRow hssfRow = (HSSFRow) rowIterator.next();
                Iterator iterator = hssfRow.cellIterator();
                List cellTempList = new ArrayList();
                
                
                //while (iterator.hasNext()) {
                for(int i=0; i < hssfRow.getLastCellNum(); i++){
                    HSSFCell hssfCell = hssfRow.getCell(i, hssfRow.CREATE_NULL_AS_BLANK);
                    hssfCell.setCellType(Cell.CELL_TYPE_STRING);
                    
                    System.out.println("Valor de celda: " + hssfCell);

                    if (hssfCell.toString().equals(null) || hssfCell.toString().equals("") || hssfCell.toString().length()<1 || hssfCell.getRichStringCellValue().getString().equals("") || hssfCell.getRichStringCellValue().getString().equals(null)) {
                        System.out.println("fila vacia");
                        //cellTempList.add(";");
                        cellTempList.add(hssfCell);
                    } else {
                        cellTempList.add(hssfCell);
                    }
                    
                   // cellTempList.add(hssfCell);
                    System.out.println("------------"+hssfCell);
                }

                cellDataList.add(cellTempList);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        escrituraArchivo.escrituraXlsToCsv(cellDataList, OutputDir);
    }
}
