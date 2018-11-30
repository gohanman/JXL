
package coop.wholefoods.jxl;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.File;
import java.io.IOException;

/**
 * Java XL data extractor
 */
public class JXL 
{
    /**
     * Fills blank cells with a string; may or may
     * not actually be needed to keep columns aligned
     */
    private String blank = "0";

    public JXL() 
    {
    }

    public void setBlankFiller(String s)
    {
        this.blank = s;
    }

    /**
     * Extract file to tab-separated
     * @param filename - input excel file
     * @param outputDir - output directory for TSVs
     */
    public void extractFile(String filename, String outputDir) throws IOException, InvalidFormatException
    {
        System.out.println( "Opening file" );
        Workbook workbook = WorkbookFactory.create(new File(filename));
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
        for (Sheet sheet: workbook) {
            this.extractSheet(sheet, outputDir);
        }
    }

    /**
     * Extract one worksheet to the output directory
     * File will have the worksheet name plus .tsv
     * @param sheet - worksheet object
     * @param outputDir - output directory
     */
    private void extractSheet(Sheet sheet, String outputDir) throws IOException, InvalidFormatException
    {
        String outfile = outputDir + sheet.getSheetName() + ".tsv";
        DataFormatter formatter = new DataFormatter();
        BufferedWriter writer = new BufferedWriter(new FileWriter(outfile));
        System.out.println("=> " + sheet.getSheetName() + " to " + outfile);
        int width = this.getSheetWidth(sheet);
        System.out.println("=> Max cells/row is " + width);
        for (Row row: sheet) {
            this.writeRow(row, width, writer, formatter); 
        }
    }

    /**
     * Scan through a sheet to find the widest row
     * @param sheet - worksheet
     * @return int - largest number of columns
     */
    private int getSheetWidth(Sheet sheet)
    {
        int width = 0;
        for (Row row: sheet) {
            int cur = row.getLastCellNum();
            if (cur > width) {
                width = cur;
            }
        }

        return width;
    }

    /**
     * Extract row data and write it to the current file
     * Rows that contain no data are ignored.
     * @param Row - row data
     * @param width - number of cells in widest row
     * @param writer - open file handle
     * @param formatter - Excel data format handler
     * @return boolean - success or failure
     */
    private boolean writeRow(Row row, int width, BufferedWriter writer, DataFormatter formatter) throws IOException, InvalidFormatException
    {
        boolean hasData = false;
        StringBuilder line = new StringBuilder();
        /**
         * The row object won't contain blanks; Loop the full width to make
         * sure rows are all the same size
         *
         * I'm not sure if "Unstale" is part of the particular file I'm
         * working with or a feature of POI but I encounter a lot of rows
         * with no data other than "Unstale" so I treat them as non-data 
         * when determining if the row is empty
         */
        for (int i=0; i<width; i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                line.append(this.blank + "\t");
            } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                switch(cell.getCachedFormulaResultType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        line.append(cell.getNumericCellValue() + "\t");
                        if (cell.getNumericCellValue() != 0) {
                            hasData = true;
                        }
                        break;
                    case Cell.CELL_TYPE_STRING:
                        String str = cell.getRichStringCellValue().getString().replace('\t', ' ');
                        if (str.length() > 0 && !str.equals("Unstale")) {
                            hasData = true;
                        }
                        line.append(str + "\t");
                        break;
                    default:
                        line.append(this.blank + "\t");
                        break;
               }
            } else {
                String cellValue = formatter.formatCellValue(cell);
                if (cellValue.length() == 0) {
                    cellValue = this.blank;
                } else {
                    cellValue = cellValue.replace('\t', ' ');
                    if (!cellValue.equals("Unstale")) {
                        hasData = true;
                    }
                }
                line.append(cellValue + "\t");
            }
        }
        line.append("\n");
        if (hasData) {
            writer.write(line.toString());
        }

        return hasData;
    }
}

