package ar.com.jengibre.filtrodla;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Charsets;
import com.google.common.base.Joiner;
import com.google.common.collect.ImmutableList;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import com.google.common.io.Files;

public class Main {

  private static final char DELIM = ';';

  public static void main(String[] args) throws Exception {
    if (args.length < 2) {
      System.out.println("pasar: <nombre filtro> <nombre planilla>");
      System.exit(1);
    }

    String filtro = args[0];
    String planilla = args[1];

    TreeSet<String> titulosMa = titulos(filtro, "ma");
    TreeSet<String> titulosDVB = titulos(filtro, "dvb");

    try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(planilla))) {
      try (BufferedWriter writer = Files.newWriter(new File("ma.csv"), Charsets.UTF_8)) {
        for (String titulo : titulosMa) {
          writer.write("\"" + titulo + "\"");
          writer.write(DELIM);
          writer.write(buscarPeli(workbook, titulo, 5 /* ma */));
          writer.write('\n');
        }
        writer.flush();
      }

      try (BufferedWriter writer = Files.newWriter(new File("dvb.csv"), Charsets.UTF_8)) {
        for (String titulo : titulosDVB) {
          writer.write("\"" + titulo + "\"");
          writer.write(DELIM);
          writer.write(buscarPeli(workbook, titulo, 6 /* dvb */));
          writer.write('\n');
        }
        writer.flush();
      }

      System.out.println("Listo");
    }
  }

  private static String buscarPeli(final XSSFWorkbook workbook, final String titulo, int columna) {
    List<Peli> pelis = Lists.newArrayList();

    ImmutableList<Sheet> reverseSheets = ImmutableList.copyOf(workbook).reverse();

    for (Sheet sheet : reverseSheets) {
      int colObservaciones = getColumn(sheet.getRow(0), "observaciones");
      int colEntregado = getColumn(sheet.getRow(0), "entregado");

      for (Row row : sheet) {
        Cell filmTitleCell = row.getCell(2); // columna C - film title
        Cell cell = row.getCell(columna); // columna ma/dbv

        cellValue(filmTitleCell).ifPresent(filmTitle -> {
          if (filmTitle.equalsIgnoreCase(titulo)) {
            String sheetName = sheet.getSheetName();
            int rowNum = row.getRowNum() + 1;
            String media = cellValue(cell).orElse("");
            String entregado = colEntregado == -1 ? "" : cellValue(row.getCell(colEntregado)).orElse("");
            String observaciones = colObservaciones == -1 ? "" : cellValue(row.getCell(colObservaciones)).orElse("");

            pelis.add(new Peli(sheetName, rowNum, media, entregado, observaciones));
          }
        });
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
        ? Optional.of(cell.getStringCellValue().trim().replaceAll("\\n", " ")) : Optional.empty();
  }

  private static TreeSet<String> titulos(String archivo, String prefijo)
      throws FileNotFoundException, IOException {
    TreeSet<String> ret = Sets.newTreeSet();

    try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(archivo))) {
      XSSFSheet sheet = wb.getSheetAt(0); // unico sheet

      for (int r = 2; r < sheet.getLastRowNum(); r++) {
        XSSFRow row = sheet.getRow(r);

        cellValue(row.getCell(0)).ifPresent(title -> {
          cellValue(row.getCell(7)).ifPresent(media -> {
            if (media.toLowerCase().startsWith(prefijo)) {
              ret.add(title);
            }
          });
        });
      }
    }

    return ret;
  }

  static class Peli {
    String sheet, ma, entregado, observaciones;

    int row;

    public Peli(String sheet, int row, String ma, String entregado, String observaciones) {
      this.sheet = sheet;
      this.row = row;
      this.ma = ma;
      this.entregado = entregado;
      this.observaciones = observaciones;
    }

    @Override
    public String toString() {
      return sheet + DELIM + row + DELIM + ma + DELIM + entregado + DELIM + observaciones;
    }
  }
}