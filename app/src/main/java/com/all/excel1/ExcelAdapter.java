package com.all.excel1;

import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;

import androidx.annotation.NonNull;
import androidx.recyclerview.widget.RecyclerView;

import java.util.ArrayList;

public class ExcelAdapter extends RecyclerView.Adapter<ExcelHolder> {
    private final ArrayList<ExcelRow> list;

    public ExcelAdapter(@NonNull ArrayList<ExcelRow> list) {
        this.list = list;
    }

    @NonNull
    @Override
    public ExcelHolder onCreateViewHolder(@NonNull ViewGroup parent, int viewType) {
        View view = LayoutInflater.from(parent.getContext()).inflate(R.layout.template1_item, parent, false);
        return new ExcelHolder(view);
    }

    @Override
    public void onBindViewHolder(@NonNull ExcelHolder holder, int position) {
        holder.onBindViewHolder(list.get(position));
    }

    @Override
    public int getItemCount() {
        return list.size();
    }
}