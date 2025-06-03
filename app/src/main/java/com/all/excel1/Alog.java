package com.all.excel1;

import static com.all.excel1.SysUtils.isUiThread;

import android.content.Context;
import android.os.Handler;
import android.util.Log;
import android.widget.TextView;
import android.widget.Toast;

import androidx.annotation.NonNull;

/**
 * Logger used in TestExcel1 - by Dennis Lang  landenlabs.com  May-2025
 */
public class Alog {
    public static final String TAG = "TestExcel1";


    public static void info(String msg) {
        Log.i(TAG, msg);
    }
    public static void warn(String msg) {
        Log.w(TAG, msg);
    }
    public static void error(String msg, Throwable tr) {
        Log.e(TAG, msg, tr);
    }

    public static void showInfo(@NonNull TextView statusTv, @NonNull String msg) {
        if (isUiThread()) {
            statusTv.setText(msg);
        } else {
            new Handler(statusTv.getContext().getMainLooper()).post(() -> {
                statusTv.setText(msg);
            });
        }
        Log.i(TAG, msg);
    }

    public static void showInfo(@NonNull Context context, @NonNull String msg) {
        if (isUiThread()) {
            Toast.makeText(context, msg, Toast.LENGTH_LONG).show();
        } else {
            new Handler(context.getMainLooper()).post(() -> {
                Toast.makeText(context, msg, Toast.LENGTH_LONG).show();
            });
        }
        Log.i(TAG, msg);
    }

    public static void showError(@NonNull TextView statusTv, @NonNull String msg, @NonNull Throwable tr) {
        if (isUiThread()) {
            statusTv.setText(msg);
        } else {
            new Handler(statusTv.getContext().getMainLooper()).post(() -> {
                statusTv.setText(msg);
            });
        }
        Log.e(TAG, msg, tr);
    }

    public static void showError(@NonNull Context context, @NonNull String msg, @NonNull Throwable tr) {
        new Handler(context.getMainLooper()).post(() -> {
            Toast.makeText(context, msg + tr, Toast.LENGTH_LONG).show();
        });
        Log.e(TAG, msg, tr);
    }

}
