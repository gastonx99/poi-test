package se.dandel.test.poi;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.OperandResolver;
import org.apache.poi.ss.formula.eval.StringEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.udf.AggregatingUDFFinder;
import org.apache.poi.ss.formula.udf.DefaultUDFFinder;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelToCsv {

    private InputStream is;
    private OutputStream os;
    private DataFormatter formatter;
    private FormulaEvaluator evaluator;

    public ExcelToCsv(InputStream is, OutputStream os) {
        this.is = is;
        this.os = os;
    }

    public void toCsv(int index) throws Exception {
        Workbook workbook = WorkbookFactory.create(is);
        addToolpack(workbook);
        this.formatter = new DataFormatter(true);
        this.evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        Sheet sheet = workbook.getSheetAt(index);
        int lastRowNum = sheet.getLastRowNum();
        for (int i1 = 0; i1 <= lastRowNum && i1 < 330; i1++) {
            List<String> rowToCsv = rowToCsv(sheet.getRow(i1));
            if ((i1 % 100) == 0) {
                System.out.println("Row " + i1 + "\t" + rowToCsv);
            }
        }
    }

    private void addToolpack(Workbook workbook) {
        String[] functionNames = { "PersonnummerBeraknaKontrollsiffra" };
        FreeRefFunction[] functionImpls = { new CalculatePersorgnrCheckDigit() };

        UDFFinder udfs = new DefaultUDFFinder(functionNames, functionImpls);
        UDFFinder udfToolpack = new AggregatingUDFFinder(udfs);
        workbook.addToolPack(udfToolpack);
    }

    public void readAndWrite() throws InvalidFormatException, IOException {
        Workbook workbook = WorkbookFactory.create(is);
        this.formatter = new DataFormatter(true);
        this.evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        Sheet sheet = workbook.getSheet("poi");
        int lastRowNum = sheet.getLastRowNum();
        for (int i1 = 0; i1 <= lastRowNum; i1++) {
            List<String> rowToCsv = rowToCsv(sheet.getRow(i1));
            System.out.println("Row " + i1 + "\t" + rowToCsv);
        }
    }

    private List<String> rowToCsv(Row row) {
        Cell cell = null;
        int lastCellNum = 0;
        List<String> csvLine = new ArrayList<String>();

        // Check to ensure that a row was recovered from the sheet as it is
        // possible that one or more rows between other populated rows could be
        // missing - blank. If the row does contain cells then...
        if (row != null) {

            // Get the index for the right most cell on the row and then
            // step along the row from left to right recovering the contents
            // of each cell, converting that into a formatted String and
            // then storing the String into the csvLine ArrayList.
            lastCellNum = row.getLastCellNum();
            for (int i = 0; i <= lastCellNum; i++) {
                cell = row.getCell(i);
                if (cell == null) {
                    csvLine.add("");
                } else {
                    if (cell.getCellType() != Cell.CELL_TYPE_FORMULA) {
                        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                            System.out.println(cell.getCellStyle().getDataFormat() + "\t"
                                    + cell.getCellStyle().getDataFormatString());
                            csvLine.add("*" + this.formatter.formatCellValue(cell));
                        } else {
                            csvLine.add(this.formatter.formatCellValue(cell));
                        }
                    } else {
                        csvLine.add(this.formatter.formatCellValue(cell, this.evaluator));
                    }
                }
            }
        }
        return csvLine;
    }

    public static class CalculatePersorgnrCheckDigit implements FreeRefFunction {

        @Override
        public ValueEval evaluate(ValueEval[] args, OperationEvaluationContext ec) {
            if (args.length != 1) {
                return ErrorEval.VALUE_INVALID;
            }
            try {
                ValueEval v1 = OperandResolver.getSingleValue(args[0], ec.getRowIndex(), ec.getColumnIndex());

                String persorgnr = OperandResolver.coerceValueToString(v1);
                return new StringEval(String.valueOf(PersorgnrUtil.calculateCheckDigit(persorgnr)));

            } catch (EvaluationException e) {
                e.printStackTrace();
                return e.getErrorEval();
            }
        }

    }
}