package cl.powerdata.conversor.escritura.archivo;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;

/**
 * Clase para escritura del contenido de archivo excel a csv.
 *
 * @author Carlos Dreves N
 * @version 1.0.0 12-12-2012.
 * @since JDK5.0
 */
public class EscrituraArchivo {

    /**
     * M�todo principal de ejecuci�n.
     *
     * @param fileName Nombre de archivo excel a convertir.
     * @param numeroHoja N�mero de Hoja de archivo excel a convertir.
     *
     *
     */
    public void escrituraXlsxToCsv(List cellDataList, String OutputDir) {
        String sFichero = OutputDir;
        BufferedWriter bw = null;
        try {
            bw = new BufferedWriter(new FileWriter(sFichero));
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        for (int i = 0; i < cellDataList.size(); i++) {

            List cellTempList = (List) cellDataList.get(i);
            for (int j = 0; j < cellTempList.size(); j++) {

                // HSSFCell hssfCell = (HSSFCell) cellTempList.get(j);
                XSSFCell xssfCell = (XSSFCell) cellTempList.get(j);
                // String stringCellValue = hssfCell.toString();
                String stringCellValue = xssfCell.toString() + ";";
                System.out.print(stringCellValue + "\t");

                try {
                    bw.write(stringCellValue + "\t");
                } catch (IOException e) {
                    e.printStackTrace();
                }

            }
            try {
                bw.write("\n");
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            System.out.println();
        }
        try {
            bw.close();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    public void escrituraXlsToCsv(List cellDataList, String OutputDir) {

        String sFichero = OutputDir;
        BufferedWriter bw = null;
        try {
            bw = new BufferedWriter(new FileWriter(sFichero));
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        for (int i = 0; i < cellDataList.size(); i++) {

            List cellTempList = (List) cellDataList.get(i);
            for (int j = 0; j < cellTempList.size(); j++) {

                HSSFCell hssfCell = (HSSFCell) cellTempList.get(j);

                //String stringCellValue = hssfCell.toString();
                String stringCellValue = hssfCell.toString() + ";";

                System.out.print(stringCellValue + "\t");

                try {
                    bw.write(stringCellValue + "\t");
                } catch (IOException e) {
                    e.printStackTrace();
                }

            }
            try {
                bw.write("\n");
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            System.out.println();
        }
        try {
            bw.close();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
}
