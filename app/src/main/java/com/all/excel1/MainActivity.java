package com.all.excel1;

import static android.view.View.GONE;
import static android.view.View.VISIBLE;
import static com.all.excel1.Alog.showInfo;
import static com.all.excel1.ExcelXSS.getFile;

import android.annotation.SuppressLint;
import android.content.Context;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.view.View;
import android.widget.TextView;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.annotation.AnyThread;
import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;
import androidx.recyclerview.widget.LinearLayoutManager;
import androidx.recyclerview.widget.RecyclerView;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * <pre>
 * Refactored and updated by -
 *   Dennis Lang - www.landenlabs.com
 *
 * Original author/source -
 *   https://github.com/yangxiaoge/AndroidExcelReadWrite
 * /pre>
 */
public class MainActivity extends AppCompatActivity {

    private final ArrayList<ExcelRow> readExcelList = new ArrayList<>();
    private Context mContext;
    private ExcelAdapter excelAdapter;
    private TextView statusTv;
    private View progress;
    private final ExcelXSS excelUtil = new ExcelXSS();

    // Must register before onCreate
    public ActivityResultLauncher<Intent> openFileLauncher = registerForActivityResult(
            new ActivityResultContracts.StartActivityForResult(),
            result -> {
                Uri uri = result.getData() != null ? result.getData().getData() : null;
                importExcelAsync(this, uri);
            });

    ActivityResultLauncher<Intent> createFileLauncher = registerForActivityResult(
            new ActivityResultContracts.StartActivityForResult(),
            result -> {
                Uri uri = result.getData() != null ? result.getData().getData() : null;
                exportExcelAsync(this, uri);
            });

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        mContext = this;
        initViews();
    }

    private void initViews() {
        statusTv = findViewById(R.id.footer);
        progress = findViewById(R.id.progress);
        RecyclerView recyclerView = findViewById(R.id.excel_content_rv);
        excelAdapter = new ExcelAdapter(readExcelList);
        recyclerView.setLayoutManager(new LinearLayoutManager(this));
        recyclerView.setNestedScrollingEnabled(false);
        recyclerView.setAdapter(excelAdapter);
    }

    @SuppressLint("NotifyDataSetChanged")
    public void onClick(View view) {
        int id = view.getId();
        findViewById(R.id.notice).setVisibility(GONE);

        if (id == R.id.import_excel_btn) {
            Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
            intent.setType("application/*");
            openFileLauncher.launch(intent);
        } else if (id == R.id.export_excel_btn && !readExcelList.isEmpty()) {
            Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
            intent.setType("application/*");
            intent.putExtra(Intent.EXTRA_TITLE, System.currentTimeMillis() + ".xlsx");
            createFileLauncher.launch(intent);
        } else if (id == R.id.clear_excel_btn) {
            if (readExcelList != null) {
                readExcelList.clear();
                excelAdapter.notifyDataSetChanged();
            }
        } else {
            showInfo(statusTv, "Please import excel before exporting");
        }
    }


    @SuppressLint("NotifyDataSetChanged")
    @AnyThread
    private void importExcelAsync(@NonNull Context context, @Nullable Uri uri) {
        if (uri != null) {
            File file = getFile(context, uri);
            showInfo(statusTv, "Importing..." + file.getName() + " size=" + file.length());
            progress.setVisibility(VISIBLE);

            // Perform work in background thread.
            new Thread(() -> {

                int sheetIdx = 0;
                List<ExcelRow> readExcelNew = excelUtil.readSheet(mContext, uri, file.getName(), sheetIdx);

                if (readExcelNew != null && !readExcelNew.isEmpty()) {
                    // readExcelList.clear();   // Allow multiple imports to append
                    readExcelList.addAll(readExcelNew);
                    showInfo(statusTv, "Successfully imported Sheets=" + excelUtil.inSheetCnt + " Rows=" + excelUtil.inRowCnt + " Size=" + file.length());
                } else if (readExcelList != null) {
                    readExcelList.clear();
                    showInfo(statusTv, "No data");
                }

                runOnUiThread(() -> {
                    progress.setVisibility(GONE);
                    excelAdapter.notifyDataSetChanged();
                }) ;
            }).start();
        } else {
            progress.setVisibility(GONE);
            showInfo(mContext, "Nothing to import");
        }
    }

    private void exportExcelAsync(@NonNull Context context, @Nullable Uri uri) {
        if (uri != null) {
            File file = getFile(context, uri);
            showInfo(statusTv, "Exporting..." + file.getName());
            progress.setVisibility(VISIBLE);

            // Perform work in background thread.
            new Thread(() -> {
                excelUtil.writeSheet(mContext, excelUtil.inSheet, excelUtil.inRows, uri, excelUtil.inSheetName);
                runOnUiThread(() ->  progress.setVisibility(GONE));
                showInfo(statusTv, "Exported  Sheets=" + excelUtil.outSheetCnt + " Rows=" + excelUtil.outRowCnt + " Size=" + file.length());
            }).start();
        } else {
            progress.setVisibility(GONE);
            showInfo(statusTv, "Nothing to import");
        }
    }
}
