package org.Minhnph;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

public class Equation {
    public final String rawA, rawB, rawC;    // chuỗi gốc
    public final double a, b, c;
    public double delta = Double.NaN, x1 = Double.NaN, x2 = Double.NaN;
    public String message;
    public final boolean valid;

    private Equation(String rawA, String rawB, String rawC,
                     double a, double b, double c, boolean valid) {
        this.rawA = rawA; this.rawB = rawB; this.rawC = rawC;
        this.a = a; this.b = b; this.c = c;
        this.valid = valid;
    }

    public static Equation fromRow(Row row, DataFormatter fmt, FormulaEvaluator eval) {
        String sa = fmt.formatCellValue(row.getCell(0), eval).trim();
        String sb = fmt.formatCellValue(row.getCell(1), eval).trim();
        String sc = fmt.formatCellValue(row.getCell(2), eval).trim();
        try {
            double a = Double.parseDouble(sa);
             double b = Double.parseDouble(sb);
            double c = Double.parseDouble(sc);
            boolean status = checkPreConditon(a,b,c);

            if(!status){
                   Equation eq = new Equation(sa, sb, sc,
                           Double.NaN, Double.NaN, Double.NaN,
                           false);
                   eq.message = "Input Invalid ";
                 return eq;       }
            return new Equation(sa, sb, sc, a, b, c, status);
        } catch (NumberFormatException ex) {
            Equation eq = new Equation(sa, sb, sc,
                    Double.NaN, Double.NaN, Double.NaN,
                    false);
            eq.message = "Input Invalid ";
            return eq;
        }
    }
           private static boolean checkPreConditon(double a, double b, double c){
                    if(a<=0 || a> 65535 || b<0 || b> 65535 || c<0 || c> 65535)  {
                        return false;
                    }
                    return true;
           }

    public void compute() {
        if (!valid) return;
        delta = b * b - 4 * a * c;
        if (a == 0) {
            if (b != 0) {
                x1 = -c / b; message = "Successfull - Phương trình bậc nhất";
            } else {
                message = (c == 0) ? "Successfully-Vô số nghiệm" : "Successfully-Vô nghiệm";
            }
        } else if (delta > 0) {
            x1 = (-b + Math.sqrt(delta)) / (2 * a);
            x2 = (-b - Math.sqrt(delta)) / (2 * a);
            message = "Successfully - 2 nghiệm thực phân biệt";
        } else if (delta == 0) {
            x1 = -b / (2 * a); message = "Successfully-Nghiệm kép";
        } else {
            message = "Successfully-2 nghiệm phức";
        }
    }
}
