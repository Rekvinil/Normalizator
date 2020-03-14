package com.normalizator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        System.out.println("Укажите путь к файлу:");
        Scanner sc = new Scanner(System.in);
        String fileName = sc.nextLine();
        File fileIn = new File(fileName);
        String fileOut = fileIn.getParent()+ "\\normalize_"+ fileIn.getName();
        System.out.println("Новый файл будет создан здеся:");
        System.out.println(fileOut);
        normalize(fileIn, fileOut);
    }

    public static XSSFWorkbook readFile(File file) {
        try {
            return new XSSFWorkbook((new FileInputStream(file)));
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public static void normalize(File file, String outFileName) throws NullPointerException {
        double max;
        double min;
        XSSFWorkbook wb = readFile(file);
        if (wb == null) {
            return;
        }
        XSSFSheet sheet = wb.getSheetAt(0);
        for (Row row : sheet) {
            Cell title = row.getCell(0);
            if (!title.getStringCellValue().equals("Name") && !title.getStringCellValue().equals("Type")) {
                max = row.getCell(1).getNumericCellValue();
                min = row.getCell(1).getNumericCellValue();
                for (Cell cell : row) {
                    if (!cell.getCellTypeEnum().equals(CellType.STRING)) {
                        double x = cell.getNumericCellValue();
                        if (x > max) max = x;
                        if (x < min) min = x;
                    }
                }
                for (Cell cell : row) {
                    if (!cell.getCellTypeEnum().equals(CellType.STRING)) {
                        double x = cell.getNumericCellValue();
                        double res = normalizeFunction(max, min, x);
                        cell.setCellValue(res);
                    }
                }
            }
        }
        writeFile(wb,outFileName);
    }

    public static double normalizeFunction(double max, double min, double x) {
        return (x-min)/(max-min);
    }

    public static void writeFile(XSSFWorkbook wb, String fileName) {
        try {
            FileOutputStream fileOut = new FileOutputStream(fileName);
            wb.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
