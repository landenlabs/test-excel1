package com.all.excel1;

import android.content.Context;
import android.view.View;
import android.view.ViewGroup;
import android.widget.LinearLayout;
import android.widget.TextView;

import androidx.annotation.NonNull;
import androidx.recyclerview.widget.RecyclerView;

import java.util.Map;

public class ExcelHolder extends RecyclerView.ViewHolder {

    private static final int CELL_PAD_PX = 4;
    private static final int CELL_MIN_WIDTH_PX = 100;
    private final LinearLayout.LayoutParams layoutParams
            = new LinearLayout.LayoutParams(
            LinearLayout.LayoutParams.WRAP_CONTENT,
            LinearLayout.LayoutParams.MATCH_PARENT);


    public ExcelHolder(@NonNull View itemView) {
        super(itemView);
    }

    public void onBindViewHolder(ExcelRow rowItem) {
        Context context = itemView.getContext();
        Map<Integer, Object> columns = rowItem.data;
        ViewGroup holder = (ViewGroup) itemView;
        holder.removeAllViews();

        // TODO - replace LinearLayout with a GridLayout to force tabular behavior.
        for (int colIdx = 0; colIdx < columns.size(); colIdx++) {
            TextView textView = new TextView(context);
            textView.setTextSize(13);
            textView.setText(String.valueOf(columns.get(colIdx)));
            textView.setLayoutParams(layoutParams);
            textView.setMinWidth(CELL_MIN_WIDTH_PX);
            textView.setPadding(CELL_PAD_PX, CELL_PAD_PX, CELL_PAD_PX, CELL_PAD_PX);
            textView.setBackgroundResource(R.drawable.textview_border);
            holder.addView(textView);
        }
    }
}
