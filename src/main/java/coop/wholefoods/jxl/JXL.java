
package coop.wholefoods.jxl;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.File;
import java.io.IOException;

public class JXL 
{
    private String blank = "0";

    public JXL() 
    {
    }

    public void setBlankFiller(String s)
    {
        this.blank = s;
    }

    public void extractFile(String filename, String outputDir) throws IOException, InvalidFormatException
    {
        System.out.println( "Opening file" );
        Workbook workbook = WorkbookFactory.create(new File(filename));
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
        for (Sheet sheet: workbook) {
            this.extractSheet(sheet, outputDir);
        }
    }

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

    private boolean writeRow(Row row, int width, BufferedWriter writer, DataFormatter formatter) throws IOException, InvalidFormatException
    {
        boolean hasData = false;
        StringBuilder line = new StringBuilder();
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

