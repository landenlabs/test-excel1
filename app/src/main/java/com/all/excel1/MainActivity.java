package com.all.excel1;

import static android.view.View.GONE;

import android.annotation.SuppressLint;
import android.content.Context;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.Bundle;
import android.provider.OpenableColumns;
import android.view.View;
import android.widget.Toast;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.annotation.AnyThread;
import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;
import androidx.recyclerview.widget.LinearLayoutManager;
import androidx.recyclerview.widget.RecyclerView;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

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

    // Must register before onCreate
    public ActivityResultLauncher<Intent> openFileLauncher = registerForActivityResult(
            new ActivityResultContracts.StartActivityForResult(),
            result -> {
                Uri uri = result.getData() != null ? result.getData().getData() : null;
                importExcelAsync(uri);
            });

    ActivityResultLauncher<Intent> createFileLauncher = registerForActivityResult(
            new ActivityResultContracts.StartActivityForResult(),
            result -> {
                Uri uri = result.getData() != null ? result.getData().getData() : null;
                exportExcelAsync(uri);
            });

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        mContext = this;
        initViews();
    }

    private void initViews() {
        RecyclerView recyclerView = findViewById(R.id.excel_content_rv);
        excelAdapter = new ExcelAdapter(readExcelList);
        recyclerView.setLayoutManager(new LinearLayoutManager(this));
        recyclerView.setNestedScrollingEnabled(false);
        recyclerView.setAdapter(excelAdapter);
    }

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
            Toast.makeText(mContext, "Please import excel before exporting", Toast.LENGTH_SHORT).show();
        }
    }

    // Helper to extract real filename from uri provided by activity/intent response.
    public String getFileName(@NonNull Uri uri) {
        String result = null;
        if (Objects.equals(uri.getScheme(), "content")) {
            try ( Cursor cursor = getContentResolver()
                    .query(uri, null, null, null, null) ) {
                if (cursor != null && cursor.moveToFirst()) {
                    int colIdx = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME);
                    result = cursor.getString(colIdx);
                }
            }
        }
        if (result == null) {
            result = uri.getPath();
            int cut = result.lastIndexOf('/');
            if (cut != -1) {
                result = result.substring(cut + 1);
            }
        }
        return result;
    }

    @SuppressLint("NotifyDataSetChanged")
    @AnyThread
    private void importExcelAsync(@Nullable Uri uri) {
        if (uri != null) {
            String filename = getFileName(uri);
            Toast.makeText(mContext, "Importing..." + filename, Toast.LENGTH_SHORT).show();

            // Perform work in background thread.
            new Thread(() -> {
                ExcelUtil excelUtil = new ExcelUtil();
                List<ExcelRow> readExcelNew = excelUtil.readExcel(mContext, uri, filename);

                if (readExcelNew != null && !readExcelNew.isEmpty()) {
                    // readExcelList.clear();
                    readExcelList.addAll(readExcelNew);
                    runOnUiThread(() -> Toast.makeText(mContext, "Successfully imported", Toast.LENGTH_SHORT).show());
                } else if (readExcelList != null) {
                    readExcelList.clear();
                    runOnUiThread(() -> Toast.makeText(mContext, "No data", Toast.LENGTH_SHORT).show());
                }

                runOnUiThread(() -> excelAdapter.notifyDataSetChanged());
            }).start();
        } else {
            Toast.makeText(mContext, "Nothing to import", Toast.LENGTH_SHORT).show();
        }
    }

    private void exportExcelAsync(@Nullable Uri uri) {
        if (uri != null) {
            String filename = getFileName(uri);
            Toast.makeText(mContext, "Exporting..." + filename, Toast.LENGTH_SHORT).show();

            // Perform work in background thread.
            new Thread(() -> {
                ExcelUtil excelUtil = new ExcelUtil();
                excelUtil.writeExcel(mContext, readExcelList, uri, filename);
            }).start();
        } else {
            Toast.makeText(mContext, "Nothing to import", Toast.LENGTH_SHORT).show();
        }
    }
}
