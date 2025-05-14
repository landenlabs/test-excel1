package com.all.excel1;

import static com.all.excel1.ExcelXSS.getCellFormatValue;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel data row extracted from .xlsx or .xlsm,  used in TestExcel1 - by Dennis Lang  landenlabs.com  May-2025
 */
public class ExcelRow {
    public final ArrayList<Cell> cells;
    public CellStyle style;
    public short height;
    public int rowNum;
    public List<String> spans;

    public ExcelRow(int capacity) {
        cells = new ArrayList<>(capacity);
    }

    public String getRaw(int colIdx) {
        Cell cell = colIdx < cells.size() ? cells.get(colIdx) : null;
        if (cell == null || cell.getColumnIndex() != colIdx) {
            for (Cell item : cells) {
                if (item != null && item.getColumnIndex() == colIdx) {
                    cell = item;
                    break;
                }
            }

        }

        if (cell != null) {
            return (cell.getCellType() == CellType.FORMULA) ? "=" + cell.getCellFormula() : getCellFormatValue(cell).toString();
        }
        return "";
    }
}
