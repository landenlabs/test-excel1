package com.all.excel1;

import android.content.Context;
import android.net.Uri;
import android.util.Log;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.annotation.WorkerThread;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.io.OutputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * <pre>
 * Refactored and updated by -
 *   Dennis Lang - www.landenlabs.com
 *
 * Original author/source -
 *   https://github.com/yangxiaoge/AndroidExcelReadWrite
 * /pre>
 */
public class ExcelUtil {
    private static final String TAG = ExcelUtil.class.getSimpleName();

    public int rowIdx = 0;
    public Exception lastEx = null;
    public List<ExcelRow> list = null;

    private static Object getCellFormatValue(Cell cell) {
        Object cellValue;
        if (cell != null) {
            switch (cell.getCellType()) {
                case BOOLEAN:
                    cellValue = cell.getBooleanCellValue();
                    break;
                case NUMERIC: {
                    cellValue = cell.getNumericCellValue();
                    break;
                }
                case FORMULA: {
                    try {
                        // Determine if the cell is in date format
                        if (DateUtil.isCellDateFormatted(cell)) {
                            // Convert to date format YYYY-mm-dd
                            cellValue = cell.getDateCellValue();
                        } else {
                            cellValue = cell.getNumericCellValue();
                        }
                    } catch (Exception ex) {
                        return cell.getStringCellValue();
                    }
                    break;
                }
                case STRING: {
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                default:
                    cellValue = "";
            }
        } else {
            cellValue = "";
        }
        return cellValue;
    }

    @Nullable
    @WorkerThread
    public List<ExcelRow> readExcel(Context context, Uri uri, String filePath) {
        list = new ArrayList<>();
        lastEx = null;
        rowIdx = 0;

        if (uri == null || filePath == null) {
            return null;
        }

        if (!filePath.contains(".xls")) {    // .xls, .xlsx, .xlsm
            Log.e(TAG, "Please select the correct Excel file");
            return null;
        }

        String extString = filePath.substring(filePath.lastIndexOf("."));
        if (".xls".equals(extString)) {
            try (InputStream is = context.getContentResolver().openInputStream(uri); Workbook wb = new HSSFWorkbook(is)) {  // .xls
                list = loadSheet(wb, 0);
            } catch (Exception ex) {
                lastEx = ex;
                Log.e(TAG, "Read excel error on row=" + rowIdx, ex);
            }
        } else {
            try (InputStream is = context.getContentResolver().openInputStream(uri); Workbook wb = new XSSFWorkbook(is)) {  // .xlsx, .xlsm
                list = loadSheet(wb, 0);
            } catch (Exception ex) {
                lastEx = ex;
                Log.e(TAG, "Read excel error on row=" + rowIdx, ex);
            }
        }
        return list;
    }

    @NonNull
    public List<ExcelRow> loadSheet(@NonNull Workbook wb, int sheetIdx) {
        list.clear();
        Sheet sheet = wb.getSheetAt(sheetIdx);

        int rowCount = sheet.getPhysicalNumberOfRows();
        for (rowIdx = 0; rowIdx < rowCount; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row != null) {
                ExcelRow itemMap = new ExcelRow();
                int colCount = row.getPhysicalNumberOfCells();
                for (int colIdx = 0; colIdx < colCount; colIdx++) {
                    try {
                        Cell cell = row.getCell(colIdx);
                        Object value = getCellFormatValue(cell);
                        itemMap.data.put(cell.getColumnIndex(), value);
                    } catch (Exception ex) {
                        Log.e(TAG, "Read excel error on row=" + rowIdx + " col=" + colIdx, ex);
                        // itemMap.data.put(colIdx, ex1.toString());
                        lastEx = ex;
                    }
                }
                list.add(itemMap);
            }
        }
        return list;
    }

    public static int getWidth(@Nullable Object obj) {
        if (obj == null) {
            return 0;
        } else if (obj instanceof String) {
            return obj.toString().length();
        } else {
            return String.valueOf(obj).length();
        }
    }

    public void writeExcel(@NonNull Context context, @NonNull List<ExcelRow> excelRows, @NonNull Uri uri, String filename) {
        try {
            int headerColCnt = excelRows.get(0).data.size();
            int maxCol = 0;
            ArrayList<Integer> maxColWidth = new ArrayList<>(headerColCnt);
            for (ExcelRow row : excelRows) {
                for (Map.Entry<Integer, Object> item : row.data.entrySet()) {
                    int colIdx = item.getKey();
                    maxCol = Math.max(maxCol, colIdx);
                    Object value = item.getValue();
                    if (colIdx < maxColWidth.size())
                        maxColWidth.set(colIdx, Math.max(maxColWidth.get(colIdx), getWidth(value)));
                    else
                        maxColWidth.add(colIdx, getWidth(value));
                }
            }

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                XSSFSheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName("Sheet1"));

                for (int colIdx = 0; colIdx < maxCol; colIdx++) {
                    // Set the cell default width to n characters
                    int numChar = maxColWidth.get(colIdx);
                    sheet.setColumnWidth(colIdx, numChar * 256);
                }

                for (int rowIdx = 0; rowIdx < excelRows.size(); rowIdx++) {
                    Row row = sheet.createRow(rowIdx);
                    ExcelRow excelRow = excelRows.get(rowIdx);
                    for (Map.Entry<Integer, Object> item : excelRow.data.entrySet()) {
                        Cell cell = row.createCell(item.getKey());
                        Object value = item.getValue();
                        if (value != null) {
                            if (value instanceof String)
                                cell.setCellValue((String) item.getValue());
                            else if (value instanceof Number)
                                cell.setCellValue((Double) item.getValue());
                            else if (value instanceof RichTextString)
                                cell.setCellValue((RichTextString) item.getValue());
                            else if (value instanceof Date)
                                cell.setCellValue((Date) item.getValue());
                            else if (value instanceof LocalDate)
                                cell.setCellValue((LocalDate) item.getValue());
                        }
                    }
                }

                try (OutputStream outputStream = context.getContentResolver().openOutputStream(uri)) {
                    workbook.write(outputStream);
                    // outputStream.flush();
                    // outputStream.close();
                } catch (Exception ex) {
                    Log.e(TAG, "Export error", ex);
                }
            } catch (Exception ex) {
                Log.e(TAG, "Export error", ex);
            }
            Log.i(TAG, "Export successful");
        } catch (Exception ex) {
            Log.e(TAG, "Export error", ex);
        }
    }
}
