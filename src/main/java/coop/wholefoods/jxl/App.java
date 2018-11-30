package coop.wholefoods.jxl;

import gnu.getopt.Getopt;

public class App 
{
    public static final String SAMPLE_XLSX_FILE_PATH = "/home/andy/jxl/RP.xlsm";

    public static void main( String[] args )
    {
        for (String s: args) {
            System.out.println(s);
        }
        JXL jxl = new JXL();
        Getopt g = new Getopt("nonsense", args, "i:o:h");
        int c;
        String arg;
        String inputfile = SAMPLE_XLSX_FILE_PATH;
        String outputdir = "./";
        while ((c = g.getopt()) != -1) {
            switch(c) {
                case 'i':
                    arg = g.getOptarg();
                    inputfile = arg;
                    System.out.println("input " + inputfile);
                    break;
                case 'o':
                    arg = g.getOptarg();
                    outputdir = arg;
                    System.out.println("output " + outputdir);
                    break;
                case 'h':
                    System.out.println("jxl [-i inputfile] [-o output directory]");
                    System.exit(1);
                    break;
            }
        }

        try {
            jxl.extractFile(inputfile, outputdir);
        } catch (Exception ex) {
            System.out.println("Error extract file: " + inputfile);
            System.out.println("Detail: " + ex.toString());
        }
    }
}

