package ar.com.jengibre.filtrodla;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Iterator;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.ImmutableList;

public class MEC {

  private static final NumberFormat NF = new DecimalFormat("0.#");

  public static void run(String filtro, String plani) throws Exception {
    XSSFWorkbook peliculas = new XSSFWorkbook(new FileInputStream(plani));

    XSSFWorkbook out = new XSSFWorkbook();
    XSSFSheet outSheet = out.createSheet();

    try (XSSFWorkbook reprocesamiento = new XSSFWorkbook(new FileInputStream(filtro))) {
      XSSFSheet sheet = reprocesamiento.getSheetAt(0);
      Iterator<Row> iterator = sheet.iterator();
      iterator.next(); // salteo la primera

      while (iterator.hasNext()) {
        buscarPeli(peliculas, (XSSFRow) iterator.next(), outSheet);
      }

      System.out.println("\n\nListo");
    }

    peliculas.close();

    out.write(new FileOutputStream("out.xlsx"));
    out.close();
  }

  private static void buscarPeli(final XSSFWorkbook workbook, final XSSFRow rowBusqueda,
      final XSSFSheet outSheet) {

    ImmutableList<Sheet> reverseSheets = ImmutableList.copyOf(workbook).reverse();
    cellValue(rowBusqueda.getCell(2)).ifPresent(titulo -> {

      sheet: for (Sheet sheet : reverseSheets) {
        for (Row _row : sheet) {
          XSSFRow row = (XSSFRow) _row;
          Optional<String> _filmTitle = cellValue(row.getCell(2)); // columna C - film title

          if (_filmTitle.isPresent()) {
            String filmTitle = _filmTitle.get();
            if (filmTitle.equalsIgnoreCase(titulo)) {
              String sheetName = sheet.getSheetName();

              System.out.println(
                  titulo + " encontrada en " + sheetName);

              XSSFRow newRow = outSheet.createRow(rowBusqueda.getRowNum());

              // Loop through source columns to add to new row
              for (int i = 0; i < row.getLastCellNum(); i++) {
                // Grab a copy of the old/new cell
                XSSFCell oldCell = row.getCell(i);
                XSSFCell newCell = newRow.createCell(i);

                // If the old cell is null jump to next cell
                if (oldCell == null) {
                  newCell = null;
                  continue;
                }

                XSSFCellStyle style = outSheet.getWorkbook().createCellStyle();
                if (oldCell.getCellStyle().getFillForegroundXSSFColor() == null) {
                  style.setFillForegroundColor(IndexedColors.WHITE.index);
                }
                else {
                  style.setFillForegroundColor(oldCell.getCellStyle().getFillForegroundXSSFColor());
                }
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                newCell.setCellStyle(style);

                // Set the cell data value
                switch (oldCell.getCellType()) {
                  case BLANK:
                    newCell.setBlank();
                    break;
                  case BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                  case ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                  case FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                  case NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                  case STRING:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                  case _NONE:
                    break;
                  default:
                    break;
                }
              }

              // en la N va el nombre del sheet
              newRow.createCell(13, CellType.STRING).setCellValue(sheetName);

              // en la O Comentarios Laboratorio
              newRow.createCell(14, CellType.STRING)
                  .setCellValue(cellValue(rowBusqueda.getCell(3)).orElse("**TITULO PELI NO ES NUMERO NI TEXTO**"));

              break sheet;
            }
          }
        }
      }
    });
  }

  private static Optional<String> cellValue(Cell cell) {
    if (cell == null) {
      return Optional.empty();
    }
    else {
      switch (cell.getCellType()) {
        case NUMERIC:
          return Optional.of(NF.format(cell.getNumericCellValue()));
        case STRING:
          return Optional.of(cell.getStringCellValue().trim().replaceAll("\\n", " "));
        default:
          return Optional.empty();
      }
    }
  }
}