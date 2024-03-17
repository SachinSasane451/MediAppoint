package com.excel.utility;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.excel.entity.Appointment;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class ExcelExporter {

    public static void exportToExcel(List<Appointment> appointments, OutputStream outputStream) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Appointments");
        int rowNum = 0;

        Row headerRow = sheet.createRow(rowNum++);
        String[] headers = {"ID", "Name", "Email", "Phone", "Reason", "Preferred Date"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        for (Appointment appointment : appointments) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(appointment.getId());
            row.createCell(1).setCellValue(appointment.getName());
            row.createCell(2).setCellValue(appointment.getEmail());
            row.createCell(3).setCellValue(appointment.getPhone());
            row.createCell(4).setCellValue(appointment.getReason());
            row.createCell(5).setCellValue(appointment.getPreferredDate());
        }

        workbook.write(outputStream);
        workbook.close();
    }
}
