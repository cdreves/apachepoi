package com.cdreves.conversor.utils;

import java.io.Serializable;
import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;

/**
 *
 * @author Sebastián Salazar Molina <sebasalazar@gmail.com>
 */
public class ExcelUtils implements Serializable {

    private static final Logger logger = Logger.getLogger(ExcelUtils.class);

    public static String getValue(Cell celda) {
        String resultado = "";
        try {
            if (celda != null) {
                switch (celda.getCellType()) {
                    case Cell.CELL_TYPE_BLANK:
                        resultado = "";
                        break;

                    case Cell.CELL_TYPE_BOOLEAN:
                        // Acá se puede cambiar la lógica por Sí/No  1/0
                        Boolean salidaBoleana = celda.getBooleanCellValue();
                        resultado = StringUtils.trimToEmpty(salidaBoleana.toString());
                        break;

                    case Cell.CELL_TYPE_ERROR:
                        resultado = "";
                        break;

                    case Cell.CELL_TYPE_FORMULA:
                        resultado = StringUtils.trimToEmpty(celda.getCellFormula());
                        break;

                    case Cell.CELL_TYPE_NUMERIC:
                        // Las fechas también son números si la memoria no me falla
                        Double salidaNumerica = celda.getNumericCellValue();
                        resultado = StringUtils.trimToEmpty(salidaNumerica.toString());
                        break;

                    case Cell.CELL_TYPE_STRING:
                        resultado = StringUtils.trimToEmpty(celda.getStringCellValue());
                        break;
                }
            }
        } catch (Exception e) {
            resultado = "";
            logger.error(e);
            logger.debug("Error al obtener valor", e);
        }
        return resultado;
    }
}
