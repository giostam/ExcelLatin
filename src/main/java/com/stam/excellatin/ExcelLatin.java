/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.stam.excellatin;

import com.ibm.icu.text.Transliterator;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Pattern;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author StamaterisG
 */
public class ExcelLatin {

    public static void main(String[] args) {
        int optionNumber = 0;
        String option = null;
        for (String arg : args) {
            if (arg.startsWith("-")) {
                optionNumber++;
                option = arg;
            }
        }
        
        if (args[0].equals("-h") || option > 1 || args.length == 0) {
            System.out.println("usage: ExcelLatin [option] filenameIn filenameOut columnNames...");
            System.out.println("option:");
            System.out.println("\t-L\tTo Latin");
            System.out.println("\t-G\tTo Greek");
        } else {
            String fileNameIn = args[0];
            String fileNameOut = args[1];
            List<String> columnNames = new ArrayList<>();
            System.out.println("Columns to transliterate\n---------------");
            for (int i = 2; i < args.length; i++) {
                columnNames.add(args[i]);
                System.out.println(args[i]);
            }
            System.out.println("\n");

            try {
                Workbook wb = WorkbookFactory.create(new File(fileNameIn));

                Workbook newWb = null;
                if (wb instanceof HSSFWorkbook) {
                    newWb = new HSSFWorkbook();
                } else if (wb instanceof XSSFWorkbook) {
                    newWb = new XSSFWorkbook();
                }
                FileOutputStream fileOut = new FileOutputStream(fileNameOut);
                if (newWb != null) {
                    Sheet sheetOut = newWb.createSheet();

                    Sheet sheet = wb.getSheetAt(0);

                    List<Integer> idxs = new ArrayList<>();

                    Row row = sheet.getRow(0);
                    for (Cell cell : row) {
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        if (cell.getStringCellValue() == null || cell.getStringCellValue().trim().equals("")) {
                            break;
                        }

                        if (columnNames.contains(cell.getStringCellValue())) {
                            idxs.add(cell.getColumnIndex());
                        }
                    }

                    for (Row rowIn : sheet) {
                        Row rowOut = sheetOut.createRow(rowIn.getRowNum());
                        if (rowIn.getRowNum() == 0) {
                            for (Cell cell : rowIn) {
                                cell.setCellType(Cell.CELL_TYPE_STRING);
                                Cell cellOut = rowOut.createCell(cell.getColumnIndex());
                                cellOut.setCellValue(cell.getStringCellValue());
                            }
                        } else {
                            for (Cell cell : rowIn) {
                                cell.setCellType(Cell.CELL_TYPE_STRING);
                                String cellVal = cell.getStringCellValue();
                                String cellNewVal = cellVal;
                                if (idxs.contains(cell.getColumnIndex()) && cellVal != null) {
                                    cellNewVal = toAscii(cellVal);
                                }
                                Cell cellOut = rowOut.createCell(cell.getColumnIndex());
                                cellOut.setCellValue(cellNewVal);
                            }
                        }
                    }

                    System.out.println("Finished!");

                    newWb.write(fileOut);
                    fileOut.close();
                }
            } catch (IOException | InvalidFormatException ex) {
                Logger.getLogger(ExcelLatin.class.toString()).log(Level.SEVERE, null, ex);
            }
        }

    }

    private final static Pattern diactiticalMarksPattern = Pattern.compile("\\p{InCombiningDiacriticalMarks}+");

    private static String deAccent(String str) {
        String nfdNormalizedString = Normalizer.normalize(str, Normalizer.Form.NFD);
        return diactiticalMarksPattern.matcher(nfdNormalizedString).replaceAll("");
    }

    private static final Transliterator toLatin = Transliterator.getInstance("Any-Latin");

    public static String toAscii(String str) {
        return deAccent(toLatin.transform(str));
    }
}
