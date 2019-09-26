package com.sa.xlhandler;

import jxl.*;
import java.io.File;
import java.util.*;

public class ExcelFileHandler {
    private static final String TABLE_END = "End";

    private static String[][] getTable(String xlFilePath, String sheetName, String tableName)
            throws Exception {
        String[][] table = null;
        boolean error = false;
        if (xlFilePath == null || xlFilePath.isEmpty()) {
            error = true;
        }
        if (sheetName == null || sheetName.isEmpty()) {
            error = true;
        }
        if (tableName == null || tableName.isEmpty()) {
            error = true;
        }
        if (error) return table;
        Workbook workbook = null;
        try {
            workbook = Workbook.getWorkbook(new File(xlFilePath));
        } catch (Exception e) {
            throw e;
        }
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            return table;
        }
        int startRow, startCol, endRow, endCol, ci, cj;
        Cell tableStart = sheet.findCell(tableName);
        if (tableStart == null) {
            return table;
        }
        startRow = tableStart.getRow();
        startCol = tableStart.getColumn();

        Cell tableEnd = sheet.findCell(tableName + TABLE_END);
        if (tableEnd == null) {
            return table;
        }
        endRow = tableEnd.getRow();
        endCol = tableEnd.getColumn();
        table = new String[endRow - startRow - 1][endCol - startCol - 1];
        ci = 0;
        for (int i = startRow + 1; i < endRow; i++, ci++) {
            cj = 0;
            for (int j = startCol + 1; j < endCol; j++, cj++) {
                table[ci][cj] = sheet.getCell(j, i).getContents();
            }
        }
        String tmp = "Read following rows from table:";
        for (int i = 0; i < table.length; i++) {
            tmp += "\n\t" + Arrays.toString(table[i]);
        }
        return table;
    }

    public static void main(String[] args) throws Exception {
        String currentPath = args[0];
        String currentSheetName = "Sheet1";
        String currentTableName = "test";
        String[][] table = ExcelFileHandler.getTable(currentPath, currentSheetName, currentTableName);
        List<Map> list = new ArrayList<Map>();
        String[] keys = table[0];
        for (int i = 1; i < table.length; i++) {
            Map<String, Object> map = new HashMap<>();
            int j = 0;
            for (; j < table[i].length; j++) {
                map.put(keys[j], table[i][j]);
            }
            list.add(map);
        }
        LinkedHashMap<String, String> tm = new LinkedHashMap<>();
        for (Map m : list) {
            String key = (String) m.get("Slave");
            if (!tm.containsKey(key)) {
                tm.put(key, (String) m.get("Dataset"));
            } else {
                tm.put(key, tm.get(key).trim() + "," + m.get("Dataset"));
            }
        }
        tm.forEach((k, v) -> {
            System.out.println("********************");
            System.out.println(k);
            System.out.println("********************");
            System.out.println(v);
            System.out.println();
            System.out.println();
        });
    }
}
