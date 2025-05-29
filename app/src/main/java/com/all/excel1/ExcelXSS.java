package com.all.excel1;

import static com.all.excel1.Alog.showError;
import static com.all.excel1.Alog.showInfo;

import android.content.Context;
import android.database.Cursor;
import android.net.Uri;
import android.os.Environment;
import android.provider.OpenableColumns;
import android.text.TextUtils;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.annotation.WorkerThread;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;

/**
 * Excel sheet read/writer methods for .xlsx or .xlsm,  used in TestExcel1 - by Dennis Lang  landenlabs.com  May-2025
 */
public class ExcelXSS {

    // Read/import
    public int inMaxColIdx = 0, inRowIdx = 0, inSheetCnt = 0, inRowCnt = 0;
    public String inSheetName;
    public String inSheetPrintArea;
    public XSSFSheet inSheet;
    public final List<ExcelRow> inRows = new ArrayList<>();
    public Exception lastEx = null;

    // Write/export
    public int outSheetCnt = 0, outRowCnt = 0;
    private final Map<Short, CellStyle> outStyles = new HashMap<>();


    @Nullable
    @WorkerThread
    public List<ExcelRow> readSheet(Context context, Uri uri, String filePath, int sheetIdx) {
        List<ExcelRow> list = null;
        lastEx = null;
        inRowIdx = inRowCnt = inSheetCnt = 0;

        if (uri == null || filePath == null) {
            return null;
        }

        if (!filePath.contains(".xlsx") && !filePath.contains(".xlsm")) {
            showInfo(context, "Please select the correct Excel file (.xlsx or .xlsm)");
            return null;
        }

        String extString = filePath.substring(filePath.lastIndexOf("."));
        try (InputStream is = context.getContentResolver().openInputStream(uri); XSSFWorkbook wb = new XSSFWorkbook(is)) {  // .xlsx, .xlsm
            list = readSheet(wb, sheetIdx);
        } catch (Exception ex) {
            lastEx = ex;
            showError(context,  "Read excel error on row=" + inRowIdx, ex);
        }
        return list;
    }

    private static char toChar(int aZero) {
        return (char)( 'A' + aZero);
    }

    private static String toColLetters(int colZeroBased) {
        if (colZeroBased < 26) {
            return "" + toChar(colZeroBased);
        } else {
            int letter2 = colZeroBased - 26;  // 0..25 = 'A'..'Z'
            return toColLetters(colZeroBased/26 -1) + toChar(letter2);
        }
    }

    @NonNull
    public List<ExcelRow> readSheet(@NonNull XSSFWorkbook wb, int sheetIdx) {
        inSheet = wb.getSheetAt(sheetIdx);
        int rowCount = inSheet.getPhysicalNumberOfRows();
        inRows.clear();

        inSheetPrintArea = wb.getPrintArea(sheetIdx);
        if (inSheetPrintArea != null) {
            inSheetPrintArea = inSheetPrintArea.replaceAll("[^!]*!", "");
        }
        inSheetCnt = wb.getNumberOfSheets();
        inRowCnt = rowCount;
        inSheetName = inSheet.getSheetName();
        inMaxColIdx = 0;

        for (inRowIdx = inSheet.getFirstRowNum(); inRowIdx <= inSheet.getLastRowNum(); inRowIdx++) {
            XSSFRow row = inSheet.getRow(inRowIdx);
            if (row != null) {
                int colCount = row.getPhysicalNumberOfCells();
                ExcelRow excelRow = new ExcelRow(colCount);
                excelRow.height = row.getHeight();
                excelRow.style = row.getRowStyle();
                excelRow.rowNum = inRowIdx;
                excelRow.ctRow = row.getCTRow();

                inMaxColIdx = Math.max(inMaxColIdx, row.getLastCellNum());
                for (int colNum = row.getFirstCellNum(); colNum <= row.getLastCellNum(); colNum++) {
                    try {
                        Cell cell = row.getCell(colNum);
                        // Object value = getCellFormatValue(cell);
                        // itemMap.data.put(cell.getColumnIndex(), value);
                        if (cell != null) {
                            excelRow.cells.add(cell);

                            /*
                            Object dbgVal = getCellFormatValue(cell);
                            if (dbgVal != null) {
                                String dbgValStr = dbgVal.toString();
                                if ( ! dbgValStr.isEmpty()) {
                                    String dbgMsg = String.format(Locale.US, "Cell,%s%d,%s",  toColLetters(colNum), row.getRowNum()+1, dbgVal);
                                    Alog.info(dbgMsg);
                                }
                            }
                             */

                        } else {
                            Alog.warn( "null cell at row=" + row.getRowNum() + " col=" + colNum);
                        }
                    } catch (Exception ex) {
                        Alog.error("Read excel error on row=" + inRowIdx + " col=" + colNum, ex);
                        // itemMap.data.put(colIdx, ex1.toString());
                        lastEx = ex;
                    }
                }
                inRows.add(excelRow);
            }
        }
        return inRows;
    }


    public void writeSheet(@NonNull Context context, @NonNull XSSFSheet inSheet, @NonNull List<ExcelRow> inRows, @NonNull Uri uri, String sheetName) {
        outStyles.clear();
        try {
            outSheetCnt = outRowCnt = 0;

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                XSSFSheet outSheet = workbook.createSheet(WorkbookUtil.createSafeSheetName(sheetName));

                if (inSheetPrintArea != null)
                    workbook.setPrintArea(workbook.getNumberOfSheets()-1, inSheetPrintArea);

                setSheetSettings(outSheet, inSheet);
                setSheet(outSheet, inRows);

                for (int colIdx = 0; colIdx <= inMaxColIdx; colIdx++) {
                    int widthIn256Units = inSheet.getColumnWidth(colIdx);
                    outSheet.setColumnWidth(colIdx, widthIn256Units);
                }

                editSheet(workbook, sheetName);

                outSheetCnt = workbook.getNumberOfSheets();
                outRowCnt = outSheet.getPhysicalNumberOfRows();
                showInfo(context, "Export saving Sheets=" + outSheetCnt + "  Rows=" + outRowCnt);
                try (OutputStream outputStream = context.getContentResolver().openOutputStream(uri)) {
                    workbook.write(outputStream);
                    outputStream.flush();
                } catch (Exception ex) {
                    showError(context, "Export error ", ex);
                }

                showInfo(context, "Export successful" );
            } catch (Exception ex) {
                showError(context, "Export error ", ex);
            }

        } catch (Exception ex) {
            showError(context, "Export error ", ex);
        }
    }

    private void editSheet(XSSFWorkbook workbook, String sheetName) {
        try {
            XSSFSheet sheet = workbook.getSheet(sheetName);
            setCell(sheet, 5, 'B', "15-May-2025");
            setCell(sheet, 7, 'L', "test-local-weather");
            setCell(sheet, 14, 'B', "tstLvl");
            setCell(sheet, 18, 'N', "@");
            setCell(sheet, 37, 'B', "test1");
            setCell(sheet, 36, 'P', "test1aa");
            setCell(sheet, 36, 'V', "test1bb");

            setCell(sheet, 39, 'B', "test2");
            setCell(sheet, 41, 'B', "test3");
        } catch (Exception ex) {
            Alog.error("Failed to edit sheet ", ex);
        }
    }

    private void setCell(XSSFSheet sheet, int rowOneBased, char colLetter, String value) {
        setCell(sheet, rowOneBased-1, colLetter - 'A', value);
    }
    private void setCell(XSSFSheet sheet, int rowZeroBased, int colZeroBased, String value) {
        if (sheet != null && sheet.getLastRowNum() > rowZeroBased) {
            XSSFRow row = sheet.getRow(rowZeroBased);
            if (row.getLastCellNum() > colZeroBased) {
                XSSFCell cell = row.getCell(colZeroBased);
                cell.setCellValue(value);
            }
        }
    }

    private void setSheet(@NonNull XSSFSheet outSheet, @NonNull List<ExcelRow> inRows) {
        outSheetCnt = outRowCnt = 0;
        Workbook workbook = outSheet.getWorkbook();

        for (ExcelRow inRow : inRows) {
            XSSFRow outRow = outSheet.createRow(inRow.rowNum);

            setStyle(workbook, outRow, inRow.style);
            outRow.setHeight(inRow.height);
            if (inRow.ctRow != null) {
                if (inRow.ctRow.isSetSpans()) {
                    outRow.getCTRow().setSpans(inRow.ctRow.getSpans());
                }
                if (inRow.ctRow.isSetCustomFormat())
                    Alog.info("custom format");
                if (inRow.ctRow.isSetCustomHeight())
                    outRow.getCTRow().setCustomHeight(inRow.ctRow.getCustomHeight());
                if (inRow.ctRow.isSetThickBot())
                    outRow.getCTRow().setThickBot(inRow.ctRow.getThickBot());
                if (inRow.ctRow.isSetThickTop())
                    outRow.getCTRow().setThickTop(inRow.ctRow.getThickTop());
            }

            for (Cell inCell : inRow.cells) {
                Cell outCell = outRow.createCell(inCell.getColumnIndex());

                switch (inCell.getCellType()) {
                    case _NONE:
                        continue;
                    case NUMERIC:
                        outCell.setCellValue(inCell.getNumericCellValue());
                        break;
                    case STRING:
                        outCell.setCellValue(inCell.getStringCellValue());
                        break;
                    case FORMULA:
                        String formula = inCell.getCellFormula();
                        // outCell.setCellFormula(formula);
                        outCell.setCellValue(getCellFormatValue(inCell).toString());
                        break;
                    case BLANK:
                        outCell.setBlank();
                        break;
                    case BOOLEAN:
                        outCell.setCellValue(inCell.getBooleanCellValue());
                        break;
                    case ERROR:
                        outCell.setCellErrorValue(inCell.getErrorCellValue());
                        break;
                }

                setStyle(workbook, outCell, inCell.getCellStyle());

                Comment comment = inCell.getCellComment();
                if (comment != null)
                    outCell.setCellComment(comment);

                Hyperlink inLink = inCell.getHyperlink();
                if (inLink instanceof XSSFHyperlink ) {
                    XSSFHyperlink outLink = (XSSFHyperlink)workbook.getCreationHelper().createHyperlink(inLink.getType());
                    outLink.setAddress(inLink.getAddress());
                    outLink.setLabel(inLink.getLabel());
                    outLink.setCellReference(((XSSFHyperlink) inLink).getCellRef());
                    outCell.setHyperlink(outLink);
                }

                /*
                Object value = getCellFormatValue(inCell);
                int colIdx = inCell.getColumnIndex();
                int colWidth = getWidth(value);
                if (colIdx < maxColWidth.size())
                    maxColWidth.set(colIdx, colWidth = Math.max(maxColWidth.get(colIdx), colWidth));
                else
                    maxColWidth.add(colIdx, colWidth);

                outSheet.setColumnWidth(colIdx, colWidth * 256);
                 */
            }
        }

        outSheetCnt = workbook.getNumberOfSheets();
        outRowCnt = outSheet.getPhysicalNumberOfRows();
    }

    private static int getWidth(@Nullable Object obj) {
        if (obj == null) {
            return 0;
        } else if (obj instanceof String) {
            return obj.toString().length();
        } else {
            return String.valueOf(obj).length();
        }
    }

    private void setSheetSettings(@NonNull XSSFSheet outSheet, @NonNull XSSFSheet inSheet)
    {
        outSheet.setAutobreaks(inSheet.getAutobreaks());
        outSheet.setDefaultColumnWidth(inSheet.getDefaultColumnWidth());
        outSheet.setDefaultRowHeight(inSheet.getDefaultRowHeight());
        outSheet.setDefaultRowHeightInPoints(inSheet.getDefaultRowHeightInPoints());
        outSheet.setDisplayGuts(inSheet.getDisplayGuts());
        outSheet.setFitToPage(inSheet.getFitToPage());

        outSheet.setForceFormulaRecalculation(inSheet.getForceFormulaRecalculation());

        PrintSetup inSheetPrintSetup = inSheet.getPrintSetup();
        PrintSetup outSheetPrintSetup = outSheet.getPrintSetup();

        outSheetPrintSetup.setPaperSize(inSheetPrintSetup.getPaperSize());
        outSheetPrintSetup.setScale(inSheetPrintSetup.getScale());
        outSheetPrintSetup.setPageStart(inSheetPrintSetup.getPageStart());
        outSheetPrintSetup.setFitWidth(inSheetPrintSetup.getFitWidth());
        outSheetPrintSetup.setFitHeight(inSheetPrintSetup.getFitHeight());
        outSheetPrintSetup.setLeftToRight(inSheetPrintSetup.getLeftToRight());
        outSheetPrintSetup.setLandscape(inSheetPrintSetup.getLandscape());
        outSheetPrintSetup.setValidSettings(inSheetPrintSetup.getValidSettings());
        outSheetPrintSetup.setNoColor(inSheetPrintSetup.getNoColor());
        outSheetPrintSetup.setDraft(inSheetPrintSetup.getDraft());
        outSheetPrintSetup.setNotes(inSheetPrintSetup.getNotes());
        outSheetPrintSetup.setNoOrientation(inSheetPrintSetup.getNoOrientation());
        outSheetPrintSetup.setUsePage(inSheetPrintSetup.getUsePage());
        outSheetPrintSetup.setHResolution(inSheetPrintSetup.getHResolution());
        outSheetPrintSetup.setVResolution(inSheetPrintSetup.getVResolution());
        outSheetPrintSetup.setHeaderMargin(inSheetPrintSetup.getHeaderMargin());
        outSheetPrintSetup.setFooterMargin(inSheetPrintSetup.getFooterMargin());
        outSheetPrintSetup.setCopies(inSheetPrintSetup.getCopies());

        Header inSheetHeader = inSheet.getHeader();
        Header outSheetHeader = outSheet.getHeader();
        outSheetHeader.setCenter(inSheetHeader.getCenter());
        outSheetHeader.setLeft(inSheetHeader.getLeft());
        outSheetHeader.setRight(inSheetHeader.getRight());

        Footer inSheetFooter = inSheet.getFooter();
        Footer outSheetFooter = outSheet.getFooter();
        outSheetFooter.setCenter(inSheetFooter.getCenter());
        outSheetFooter.setLeft(inSheetFooter.getLeft());
        outSheetFooter.setRight(inSheetFooter.getRight());

        outSheet.setHorizontallyCenter(inSheet.getHorizontallyCenter());
        outSheet.setMargin(PageMargin.LEFT, inSheet.getMargin(PageMargin.LEFT));
        outSheet.setMargin(PageMargin.RIGHT, inSheet.getMargin(PageMargin.RIGHT));
        outSheet.setMargin(PageMargin.TOP, inSheet.getMargin(PageMargin.TOP));
        outSheet.setMargin(PageMargin.BOTTOM, inSheet.getMargin(PageMargin.BOTTOM));

        outSheet.setPrintGridlines(inSheet.isPrintGridlines());
        outSheet.setRowSumsBelow(inSheet.getRowSumsBelow());
        outSheet.setRowSumsRight(inSheet.getRowSumsRight());
        outSheet.setVerticallyCenter(inSheet.getVerticallyCenter());
        outSheet.setDisplayFormulas(inSheet.isDisplayFormulas());
        outSheet.setDisplayGridlines(inSheet.isDisplayGridlines());
        outSheet.setDisplayRowColHeadings(inSheet.isDisplayRowColHeadings());
        outSheet.setDisplayZeros(inSheet.isDisplayZeros());
        outSheet.setPrintGridlines(inSheet.isPrintGridlines());
        outSheet.setRightToLeft(inSheet.isRightToLeft());
        outSheet.setZoom(100);
        //copyPrintTitle(outSheet, inSheet);

        if (inSheet.getNumMergedRegions() > 0) {
            for (CellRangeAddress region : inSheet.getMergedRegions()) {
                outSheet.addMergedRegion(region);
            }
        }

        for (int colIdx = 0; colIdx < inMaxColIdx; colIdx++) {
            if (inSheet.isColumnHidden(colIdx)) {
                outSheet.setColumnHidden(colIdx, true);
            }

            CellStyle style = inSheet.getColumnStyle(colIdx);
            if (style != null)
                outSheet.setDefaultColumnStyle(colIdx, style);
        }

        int[] cbreaks = inSheet.getColumnBreaks();
        if (cbreaks != null && cbreaks.length > 0) {
            Alog.info("got cbreaks");
        }
        int[] rbreaks = inSheet.getRowBreaks();
        if (rbreaks != null && cbreaks.length > 0) {
            Alog.info("got cbreaks");
        }

        CellRangeAddress range = inSheet.getDimension();
        if (range != null)
            outSheet.setDimensionOverride(range);
    }
    
    private void setStyle(@NonNull Workbook workbook, @NonNull Cell outCell, @Nullable CellStyle inStyle) {
        // Avoid making duplicate styles
        if (inStyle != null) {
            CellStyle outStyle = outStyles.get(inStyle.getIndex());
            if (outStyle == null) {
                outStyle = workbook.createCellStyle();
                outStyle.cloneStyleFrom(inStyle);
                outStyles.put(inStyle.getIndex(), outStyle);
            }

            outCell.setCellStyle(outStyle);
        }
    }

    private void setStyle(@NonNull Workbook workbook, @NonNull Row outRow, @Nullable CellStyle inStyle) {
        // Avoid making duplicate styles
        if (inStyle != null) {
            CellStyle outStyle = outStyles.get(inStyle.getIndex());
            if (outStyle == null) {
                outStyle = workbook.createCellStyle();
                outStyle.cloneStyleFrom(inStyle);
                outStyles.put(inStyle.getIndex(), outStyle);
            }

            outRow.setRowStyle(outStyle);
        }
    }


    static Object getCellFormatValue(@NonNull Cell cell) {
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

    // Helper to extract real filename from uri provided by activity/intent response.
    public static File getFile(@NonNull Context context, @NonNull Uri uri) {

        String result = null;
        if (Objects.equals(uri.getScheme(), "content")) {
            try (Cursor cursor = context.getContentResolver()
                    .query(uri, null, null, null, null)) {
                if (cursor != null && cursor.moveToFirst()) {
                    int colIdx = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME);
                    result = cursor.getString(colIdx);
                }
            }
        }
        if (result == null) {
            result = uri.getPath();
        }

        //     Environment.DIRECTORY_DOWNLOADS
        String path = uri.toString().replaceAll(".*[.]providers[.]([^.]+).*", "$1").toLowerCase(Locale.US);
        String subDir = Environment.DIRECTORY_DOWNLOADS;
        if (path.contains("download"))
            subDir = Environment.DIRECTORY_DOWNLOADS;
        else if (path.contains("picture"))
            subDir = Environment.DIRECTORY_PICTURES;
        else if (path.contains("document"))
            subDir = Environment.DIRECTORY_DOCUMENTS;
        else if (path.contains("movie"))
            subDir = Environment.DIRECTORY_MOVIES;

        File file = new File(new File(Environment.getExternalStorageDirectory(), subDir), result);
        return file;
    }

    public static String toCapitalize(final String str) {
        if (TextUtils.isEmpty(str) || str.length() == 1)
            return str;
        return Character.toUpperCase(str.charAt(0))
                + str.substring(1).toLowerCase(Locale.US);
    }
}
