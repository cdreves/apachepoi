package powerdata.conversor.procesosglobales;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

import javax.print.DocFlavor.STRING;

import org.apache.poi.hssf.usermodel.HSSFCell;
//import org.apache.poi.ss.examples.ToCSV;
import org.apache.poi.xssf.usermodel.XSSFCell;

/**
 * Clase con procesos globales del programa.
 * 
 * @author Carlos Dreves N
 * @version 1.0.0 12-12-2012.
 * @since JDK5.0
 */

public class ProcesoGlobal {

	/**
	 * Metodo para validar extension de archivo de entrada.
	 * 
	 * @param fileName
	 *            Nombre de archivo excel a procesar.
	 * @return numero representativo a cada extension. \n 0 = xls \n 1 = xls \n
	 *         -1 = Formato no valido.
	 * 
	 **/

	public int validarExtension(String fileName) {

		String extension = null;

		try {
			extension = fileName.substring(fileName.lastIndexOf("."),
					fileName.length());
			if (extension.equals(".xlsx") == true) {
				return 0; // retorno 0 es xlsx
			} else if (extension.equals(".xls") == true) {
				return 1; // retorno 1 es xls
			} else {
				return -1; // retorno -1 es xls
			}
		} catch (Exception e) {
			System.out.println("Verificar Extension archivo de entrada" + e);
			return -1;
		}

	}
}
