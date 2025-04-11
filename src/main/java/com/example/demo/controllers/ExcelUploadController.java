package com.example.demo.controllers;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import jakarta.servlet.http.HttpServletResponse;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.DateUtil;
/**
 *
 * @author raulin
 */
@Controller
public class ExcelUploadController {

    @GetMapping("/")
    public String index() {
        return "index";
    }

    @PostMapping("/upload")
    public void handleFileUpload(@RequestParam("file") MultipartFile file, HttpServletResponse response) {
        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            StringBuilder sql = new StringBuilder();

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String tableName = sheet.getSheetName().replaceAll("\\s+", "_").toUpperCase();

                Row header = sheet.getRow(0);
                if (header == null) continue;

                sql.append("BEGIN EXECUTE IMMEDIATE 'DROP TABLE \"").append(tableName).append("\"'; EXCEPTION WHEN OTHERS THEN NULL; END;\n/\n");

                sql.append("CREATE TABLE \"").append(tableName).append("\" (\n");

                // detectar tipo celda
                Map<Integer, String> columnTypes = new HashMap<>();

                int cellCount = header.getLastCellNum();
                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;

                    for (int j = 0; j < cellCount; j++) {
                        Cell cell = row.getCell(j);
                        if (cell == null) continue;

                        String detectedType = detectCellType(cell);

                        if (columnTypes.get(j) == null || columnTypes.get(j).equals("VARCHAR2(4000)")) {
                            columnTypes.put(j, detectedType);
                        }
                    }
                }
                // Crear nombre columnas
                for (int j = 0; j < cellCount; j++) {
                    String columnName = header.getCell(j).getStringCellValue()
                            .replaceAll("[^a-zA-Z0-9_]", "_").toUpperCase();

                    String columnType = columnTypes.get(j);
                    sql.append("  \"").append(columnName).append("\" ").append(columnType);
                    if (j < cellCount - 1) sql.append(",\n");
                }
                sql.append("\n);\n\n");

                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;

                    sql.append("INSERT INTO \"").append(tableName).append("\" VALUES (");
                    for (int c = 0; c < cellCount; c++) {
                        Cell cell = row.getCell(c);
                        String value = "NULL";
                        if (cell != null) {
                            switch (cell.getCellType()) {
                                case STRING:
                                    value = "'" + cell.getStringCellValue().replace("'", "''") + "'";
                                    break;
                                case NUMERIC:
                                    if (DateUtil.isCellDateFormatted(cell)) {
                                        value = "TO_DATE('" + cell.getDateCellValue() + "', 'YYYY-MM-DD')";
                                    } else {
                                        value = String.valueOf(cell.getNumericCellValue());
                                    }
                                    break;
                                case BOOLEAN:
                                    value = cell.getBooleanCellValue() ? "'1'" : "'0'";
                                    break;
                                default:
                                    value = "NULL";
                                    break;
                            }
                        }
                        sql.append(value);
                        if (c < cellCount - 1) sql.append(", ");
                    }
                    sql.append(");\n");
                }
                sql.append("\n");
            }

            response.setContentType("text/sql");
            response.setHeader("Content-Disposition", "attachment; filename=output.sql");
            response.getWriter().write(sql.toString());

        } catch (Exception e) {
            throw new RuntimeException("Failed to process file: " + e.getMessage(), e);
        }
    }

    private String detectCellType(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return "VARCHAR2(4000)";
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return "DATE";
                } else {
                    return "NUMBER";
                }
            case BOOLEAN:
                return "CHAR(1)";
            case FORMULA:
                return "VARCHAR2(4000)";
            default:
                return "VARCHAR2(4000)";
        }
    }
}