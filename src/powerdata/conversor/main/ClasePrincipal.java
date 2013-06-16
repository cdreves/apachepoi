package powerdata.conversor.main;

import java.io.File;

import powerdata.conversor.lectura.excel.*;
import powerdata.conversor.procesosglobales.ProcesoGlobal;

/**
 * Clase Principal de programa.
 *
 * @author Carlos Dreves N
 * @version 1.0.0 12-12-2012.
 * @since JDK5.0
 */
public class ClasePrincipal {

    private static LecturaInput LecturaInput = new LecturaInput();
    private static ProcesoGlobal ProcesoGlobal = new ProcesoGlobal();

    /**
     * M�todo principal de ejecuci�n.
     *
     * @param fileName Nombre de archivo excel a convertir.
     * @param numeroHoja N�mero de Hoja de archivo excel a convertir.
     *
     *
     */
    public static void main(String[] args) {

        if (args.length == 3) { // validar cantidad de parametros
            String fileName = args[0];
            Integer numeroHoja = Integer.parseInt(args[1]) - 1; // Indice de
            // Hoja menos 1
            String outputDir = args[2];
            //File.separator // separador Java

            int resultado = ProcesoGlobal.validarExtension(fileName);

            if (resultado == 0) {
                LecturaInput.leerArchivoXlsx(fileName, numeroHoja, outputDir); // Lee
                // xlsx
                // e invoca
                // a m�todo
                // para
                // crear
                // archivo
                // plano
            } else if (resultado == 1) {
                LecturaInput.leerArchivoXls(fileName, numeroHoja, outputDir); // Lee
                // xls
                // e
                // invoca a
                // m�todo
                // para
                // crear
                // archivo
                // plano
            } else {
                System.out.println("Archivo de extensi�n erronea");
            }

        }
    }
}
