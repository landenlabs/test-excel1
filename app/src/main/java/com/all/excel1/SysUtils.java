package com.all.excel1;

import android.app.Activity;
import android.content.Context;
import android.content.ContextWrapper;
import android.os.Looper;

import androidx.annotation.Nullable;

import java.util.Objects;

/**
 * System utility methods used in TestExcel1 - by Dennis Lang  landenlabs.com  May-2025
 */
public class SysUtils {

    public static boolean isUiThread() {
        // TO DO - see Context.isUiContext()
        return Objects.equals(Thread.currentThread().getId(), Looper.getMainLooper().getThread().getId());
    }

    @Nullable
    public static Activity getActivity(@Nullable Context context) {
        if (context instanceof Activity) {
            return (Activity) context;
        } else if (context == null) {
            return null;
        } else if (context instanceof ContextWrapper) {
            return getActivity(((ContextWrapper) context).getBaseContext());
        }
        return null;
    }
}
