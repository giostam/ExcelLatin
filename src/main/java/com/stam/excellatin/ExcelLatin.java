/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.stam.excellatin;

import com.ibm.icu.text.Transliterator;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Pattern;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
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

    private static final List<String> validOptions = new ArrayList<>();

    static {
        validOptions.add("-L");
        validOptions.add("-G");
        validOptions.add("-d");
    }

    public static void main(String[] args) {
        List<String> options = new ArrayList<>();
        int startIndex = 0;
        for (String arg : args) {
            if (validOptions.contains(arg)) {
                options.add(arg);
                startIndex++;
            }
        }

        if (args[0].equals("-h") || args.length < 3) {
            System.out.println("usage: ExcelLatin [options] filenameIn filenameOut columnNames...");
            System.out.println("options:");
            System.out.println("\t-L\tto Latin (default)");
            System.out.println("\t-G\tto Greek");
            System.out.println("\t-d\tdon't deaccent");
            System.out.println("\t-h\thelp");
        } else {
            boolean greekToLatin = false;
            boolean latinToGreek = false;
            Transliterator transliterator = null;
            if ((!options.contains("-L") && !options.contains("-G")) || options.contains("-L")) {
                transliterator = Transliterator.getInstance("Greek-Latin/UNGEGN");
                System.out.println("\nTransliterating Greek to Latin");
                greekToLatin = true;
            } else if (options.contains("-G")) {
                transliterator = Transliterator.getInstance("Latin-Greek/UNGEGN");
                System.out.println("\nTransliterating Latin to Greek");
                latinToGreek = true;
            }

            if (transliterator == null) {
                System.out.println("Not a valid option for the transliteration language");
                return;
            }

            boolean deAccent = true;
            if (options.contains("-d")) {
                deAccent = false;
                System.out.println("Will not deaccent");
            }

            String fileNameIn = args[startIndex];
            String fileNameOut = args[startIndex + 1];
            List<String> columnNames = new ArrayList<>();
            System.out.println("\nColumns to transliterate\n---------------------------");
            for (int i = startIndex + 2; i < args.length; i++) {
                columnNames.add(args[i]);
                System.out.println(args[i]);
            }
            System.out.println("\n");

            try {
                File file = new File(fileNameIn);
                if (!file.exists()) {
                    System.out.println("The file " + fileNameIn + " was not found");
                    return;
                }

                Map<String, String> mapTransformations = new HashMap<>();
                Scanner sc = new Scanner(new FileReader("map.txt"));
                while (sc.hasNextLine()) {
                    String greekEntry = sc.next();
                    String latinEntry = sc.next();

                    if (greekToLatin) {
                        mapTransformations.put(greekEntry, latinEntry);
                    } else if (latinToGreek) {
                        mapTransformations.put(latinEntry, greekEntry);
                    }
                }

                DataFormatter formatter = new DataFormatter();
                Workbook wb = WorkbookFactory.create(file);

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
                        String cellVal = formatter.formatCellValue(cell);
                        if (cellVal == null || cellVal.trim().equals("")) {
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
                                String cellVal = formatter.formatCellValue(cell);
                                String cellNewVal = cellVal;
                                if (idxs.contains(cell.getColumnIndex()) && cellVal != null) {
                                    if (mapTransformations.containsKey(cellVal)) {
                                        cellNewVal = mapTransformations.get(cellVal);
                                    } else {
                                        if (deAccent) {
                                            cellNewVal = deAccent(transliterator.transform(cellVal));
                                        } else {
                                            cellNewVal = transliterator.transform(cellVal);
                                        }
                                    }
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
}
