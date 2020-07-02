package ar.com.jengibre.filtrodla;

public class Main {

  private static void usage() {
    System.out.println("usar:");
    System.out.println("superflash <nombre filtro> <nombre planilla>");
    System.out.println("o");
    System.out.println("mec <nombre peliculas> <nombre reprocesamiento>");
  }

  public static void main(String[] args) throws Exception {
    if (args.length < 3) {
      usage();
    }

    if ("superflash".equals(args[0])) {
      SuperFlash.run(args[1], args[2]);
    }
    else if ("mec".equals(args[0])) {
      MEC.run(args[1], args[2]);
    }
    else {
      usage();
    }
  }
}