package ar.com.jengibre.filtrodla;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;
import java.util.Optional;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Charsets;
import com.google.common.base.Joiner;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import com.google.common.io.Files;

public class Main {

  private static final char DELIM = ';';

  public static void main(String[] args) throws Exception {
    if (args.length < 3) {
      System.out.println("pasar: <nombre filtro> <nombre planilla> <p/s>");
      System.exit(1);
    }

    String filtro = args[0];
    String planilla = args[1];
    boolean series = args[2].equalsIgnoreCase("s");

    TreeSet<String> titulos = titulos(filtro);
    XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(planilla));
    BufferedWriter writer = Files.newWriter(new File("salida.csv"), Charsets.UTF_8);

    for (String titulo : titulos) {
      writer.write("\"" + titulo + "\"");
      writer.write(DELIM);
      writer.write(series ? buscarSerie(workbook, titulo) : buscarPeli(workbook, titulo));
      writer.write('\n');
    }

    writer.flush();
    System.out.println("Listo");
  }

  private static String buscarPeli(final XSSFWorkbook workbook, final String titulo) throws Exception {
    List<Peli> pelis = Lists.newArrayList();
    for (Sheet sheet : workbook) {

      int colObservaciones = getColumn(sheet.getRow(0), "observaciones");

      for (Row row : sheet) {
        Cell filmTitleCell = row.getCell(2); // columna C - film title
        Cell maCell = row.getCell(5); // columna F - ma

        Optional<String> _filmTitle = cellValue(filmTitleCell);
        if (_filmTitle.isPresent()) {
          if (_filmTitle.get().equalsIgnoreCase(titulo)) {
            String sheetName = sheet.getSheetName();
            int rowNum = row.getRowNum() + 1;
            String ma = cellValue(maCell).orElse("");
            String obs = colObservaciones == -1 ? "" : cellValue(row.getCell(colObservaciones)).orElse("");

            pelis.add(new Peli(sheetName, rowNum, ma, obs));
          }
        }
      }
    }

    return Joiner.on(DELIM).join(pelis);
  }

  private static short getColumn(Row row, String name) {
    if (row != null) {
      for (short i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
        Cell cell = row.getCell(i);
        if (cell != null && cell.getCellTypeEnum() == CellType.STRING
            && cell.getStringCellValue().trim().equalsIgnoreCase(name)) {
          return i;
        }
      }
    }

    return -1;
  }

  private static Optional<String> cellValue(Cell cell) {
    return (cell != null && cell.getCellTypeEnum() == CellType.STRING)
        ? Optional.of(cell.getStringCellValue().trim()) : Optional.empty();
  }

  private static String buscarSerie(final XSSFWorkbook workbook, final String titulo) throws Exception {
    List<Integer> ret = Lists.newArrayList();
    Sheet sheet = workbook.getSheetAt(0);
    for (Row row : sheet) {
      Cell cell = row.getCell(0); // columna A - TITULO
      if (cell != null && cell.getCellTypeEnum() == CellType.STRING) {
        String filmTitle = cell.getStringCellValue().trim();
        if (filmTitle.equalsIgnoreCase(titulo)) {
          ret.add(row.getRowNum() + 1);
        }
      }
    }

    return Joiner.on(DELIM).join(ret);
  }

  private static TreeSet<String> titulos(String archivo) throws Exception {
    TreeSet<String> ret = Sets.newTreeSet();

    try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(archivo))) {
      XSSFSheet sheet = wb.getSheetAt(0); // unico sheet

      for (int r = 2; r < sheet.getLastRowNum(); r++) {
        XSSFRow row = sheet.getRow(r);

        XSSFCell cell = row.getCell(0); // program title
        if (cell != null && cell.getCellTypeEnum() == CellType.STRING) {
          String title = cell.getStringCellValue().trim();
          if (!title.isEmpty()) {
            ret.add(cell.getStringCellValue().trim());
          }
        }
      }
    }

    return ret;
  }

  static class Peli {
    String sheet, ma, observaciones;

    int row;

    public Peli(String sheet, int row, String ma, String observaciones) {
      this.sheet = sheet;
      this.row = row;
      this.ma = ma;
      this.observaciones = observaciones;
    }

    @Override
    public String toString() {
      return sheet + DELIM + row + DELIM + ma + DELIM + observaciones;
    }
  }
}