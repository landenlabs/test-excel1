package com.all.excel1;

import android.content.Context;
import android.graphics.Typeface;
import android.view.View;
import android.view.ViewGroup;
import android.widget.LinearLayout;
import android.widget.TextView;

import androidx.annotation.NonNull;
import androidx.recyclerview.widget.RecyclerView;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.Objects;

/**
 * RecyclerView holder used in TestExcel1 - by Dennis Lang  landenlabs.com  May-2025
 */
public class ExcelHolder extends RecyclerView.ViewHolder {

    public static final int NO_COLOR = 0x12345678;
    private static final int CELL_PAD_PX = 4;
    private static final int CELL_MIN_WIDTH_PX = 100;
    private final LinearLayout.LayoutParams layoutParams
            = new LinearLayout.LayoutParams(LinearLayout.LayoutParams.WRAP_CONTENT, LinearLayout.LayoutParams.MATCH_PARENT);

    public ExcelHolder(@NonNull View itemView) {
        super(itemView);
    }

    public static int getARGB(Color color, int noColor) {
        if (color instanceof XSSFColor xssfColor) {
            try {
                if (xssfColor.hasAlpha()) {
                    byte[] argb = xssfColor.getARGB();
                    if (argb != null)
                        return android.graphics.Color.argb(argb[0], argb[1], argb[2], argb[3]);
                }
            } catch (Throwable tr) {
                // Alog.warn(TAG, "failed to get argb ", tr);
            }
            try {
                byte[] rgb = xssfColor.getCTColor().getRgb();
                if (rgb != null)
                    return android.graphics.Color.rgb(rgb[0], rgb[1], rgb[2]);
            } catch (Throwable tr) {
                // Alog.warn(TAG, "failed to get rgb ", tr);
            }
        }
        return noColor;
    }

    @SafeVarargs
    public static <TT> boolean isAny(TT want, TT... anyOf) {
        for (TT item : anyOf) {
            if (Objects.equals(want, item))
                return true;
        }
        return false;
    }

    public void onBindViewHolder(ExcelRow rowItem) {
        Context context = itemView.getContext();
        ViewGroup holder = (ViewGroup) itemView;
        holder.removeAllViews();

        // TODO - replace LinearLayout with a GridLayout to force tabular behavior.

        for (int colIdx = 0; colIdx < rowItem.cells.size(); colIdx++) {
            Cell cell = rowItem.cells.get(colIdx);
            if (cell != null && !isAny(cell.getCellType(), CellType._NONE, CellType.BLANK, CellType.ERROR)) {

                TextView textView = new TextView(context);
                textView.setTextSize(13);

                XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getCellStyle();

                XSSFFont font = cellStyle.getFont();
                if (font.getBold())
                    textView.setTypeface(Typeface.DEFAULT_BOLD);

                int fontColor = getARGB(font.getXSSFColor(), NO_COLOR);
                // int bgColor = getARGB(cellStyle.getFillBackgroundColorColor(), NO_COLOR);
                int fgColor = getARGB(cellStyle.getFillForegroundColorColor(), NO_COLOR);
                if (fontColor != NO_COLOR)
                    textView.setTextColor(fontColor);
                if (fgColor != NO_COLOR)
                    textView.setBackgroundColor(fgColor);
                else
                    textView.setBackgroundResource(R.drawable.textview_border);

                textView.setText(rowItem.getRaw(colIdx));
                textView.setLayoutParams(layoutParams);
                textView.setMinWidth(CELL_MIN_WIDTH_PX);
                textView.setPadding(CELL_PAD_PX, CELL_PAD_PX, CELL_PAD_PX, CELL_PAD_PX);

                holder.addView(textView);
            }
        }
    }
}
