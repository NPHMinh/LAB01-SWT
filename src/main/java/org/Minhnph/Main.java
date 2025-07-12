package org.Minhnph;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {

    public static void main(String[] args) throws IOException {
        String inputPath  = "src/main/resources/phuongtrinh2.xlsx";
        String outputPath = "src/main/resources/phuongtrinh2_result.xlsx";

        // 1. Đọc file nguồn
        try (Workbook wb = WorkbookFactory.create(new FileInputStream(inputPath))) {
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
            Sheet sheet = wb.getSheetAt(0);
            List<Equation> list = new ArrayList<>();
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;  // bỏ header
                list.add(Equation.fromRow(row, fmt, eval));
            }

            // Debug: in số phương trình
            System.out.println("Đã đọc " + list.size() + " phương trình.");

            // 2. Tính toán
            list.forEach(Equation::compute);

            // In kết quả ra terminal (console)
            System.out.printf("%-10s %-10s %-10s %-12s %-15s %-15s %s%n",
                    "a", "b", "c", "delta", "x1", "x2", "message");
            for (Equation eq : list) {
                String aStr = eq.valid ? String.valueOf(eq.a) : eq.rawA;
                String bStr = eq.valid ? String.valueOf(eq.b) : eq.rawB;
                String cStr = eq.valid ? String.valueOf(eq.c) : eq.rawC;
                String deltaStr = !Double.isNaN(eq.delta) ? String.valueOf(eq.delta) : "";
                String x1Str = !Double.isNaN(eq.x1) ? String.valueOf(eq.x1) : "";
                String x2Str = !Double.isNaN(eq.x2) ? String.valueOf(eq.x2) : "";

                System.out.printf("%-10s %-10s %-10s %-12s %-15s %-15s %s%n",
                        aStr, bStr, cStr, deltaStr, x1Str, x2Str, eq.message == null ? "" : eq.message);
            }

            // 3. Ghi kết quả với SXSSFWorkbook để streaming
            try (SXSSFWorkbook outWb = new SXSSFWorkbook((XSSFWorkbook) wb, 100)) {
                // Nếu đã có sheet "KetQua", xóa trước khi tạo
                int existingSheetIdx = outWb.getSheetIndex("KetQua");
                if (existingSheetIdx != -1) {
                    outWb.removeSheetAt(existingSheetIdx);
                }

                // Tạo sheet kết quả
                Sheet outSheet = outWb.createSheet("KetQua");

                // Header
                Row header = outSheet.createRow(0);
                String[] heads = {"a", "b", "c", "delta", "x1", "x2", "message"};
                for (int i = 0; i < heads.length; i++) {
                    header.createCell(i).setCellValue(heads[i]);
                }

                // CellStyle cho ô invalid (đỏ)
                CellStyle invalidStyle = outWb.createCellStyle();
                Font redFont = outWb.createFont();
                redFont.setColor(IndexedColors.RED.getIndex());
                invalidStyle.setFont(redFont);

                // Viết dữ liệu
                for (int i = 0; i < list.size(); i++) {
                    Equation eq = list.get(i);
                    Row r = outSheet.createRow(i + 1);

                    // Ghi a, b, c (giữ nguyên chuỗi nếu invalid)
                    String[] raws = {eq.rawA, eq.rawB, eq.rawC};
                    double[] vals = {eq.a, eq.b, eq.c};
                    for (int j = 0; j < 3; j++) {
                        Cell cell = r.createCell(j);
                        if (eq.valid) {
                            cell.setCellValue(vals[j]);
                        } else {
                            cell.setCellValue(raws[j]);
                            cell.setCellStyle(invalidStyle);
                        }
                    }

                    // Ghi delta, x1, x2 (bỏ qua nếu NaN)
                    double[] results = {eq.delta, eq.x1, eq.x2};
                    for (int j = 0; j < results.length; j++) {
                        if (!Double.isNaN(results[j])) {
                            r.createCell(3 + j).setCellValue(results[j]);
                        }
                    }

                    // Ghi message
                    Cell msgCell = r.createCell(6);
                    msgCell.setCellValue(eq.message);
                    if (!eq.valid) {
                        msgCell.setCellStyle(invalidStyle);
                    }
                }

                // Xóa sheet cũ, chỉ giữ "KetQua" (tuỳ yêu cầu, nếu cần xoá thêm sheet khác)
                // Có thể không cần dòng này nếu chỉ muốn giữ "KetQua"
                int sheetCount = outWb.getNumberOfSheets();
                for (int i = sheetCount - 1; i >= 0; i--) {
                    if (!outWb.getSheetName(i).equals("KetQua")) {
                        outWb.removeSheetAt(i);
                    }
                }

                // Chuyển focus sang sheet "KetQua"
                int idx = outWb.getSheetIndex("KetQua");
                outWb.setActiveSheet(idx);
                outWb.setSelectedTab(idx);

                // Ghi file ra đĩa
                try (FileOutputStream fos = new FileOutputStream(new File(outputPath))) {
                    outWb.write(fos);
                }
                outWb.dispose();

                System.out.println("Đã ghi " + (list.size() + 1) + " dòng vào " + outputPath);
            }
        }
    }
}