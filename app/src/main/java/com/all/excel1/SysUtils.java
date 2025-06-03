package com.all.excel1;

import android.app.Activity;
import android.content.Context;
import android.content.ContextWrapper;
import android.content.Intent;
import android.net.Uri;
import android.os.Looper;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.core.content.FileProvider;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
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


    public static long copyFile(@NonNull Context context, @NonNull File src, File dst) {
        try {
            String resolvePkg = BuildConfig.APPLICATION_ID + ".provider";
            Uri srcUri = FileProvider.getUriForFile(Objects.requireNonNull(context.getApplicationContext()), resolvePkg, src);
            context.grantUriPermission(resolvePkg, srcUri, Intent.FLAG_GRANT_READ_URI_PERMISSION);
            // context.grantUriPermission(context.getPackageName(), srcUri,  Intent.FLAG_GRANT_READ_URI_PERMISSION);

            copyFile(context, srcUri, dst);
        } catch (Exception ex2) {
            Alog.showError(context, "Failed to copy file to " + dst.getName(), ex2);
        }
        return -1L;
    }

    public static long copyFile(@NonNull Context context, @NonNull Uri srcUri, File dst) {
        try {
            try (InputStream in = context.getContentResolver().openInputStream(srcUri)) {
                // try (InputStream in = new FileInputStream(src)) {
                try (OutputStream out = new FileOutputStream(dst)) {
                    // Transfer bytes from in to out
                    byte[] buf = new byte[1024];
                    int len;
                    while ((len = in.read(buf)) > 0) {
                        out.write(buf, 0, len);
                    }
                    return dst.length();
                }
            } catch (Exception ex) {
                Alog.showError(context, "Failed to copy file to " + dst.getName(), ex);
            }
        } catch (Exception ex2) {
            Alog.showError(context, "Failed to copy file to " + dst.getName(), ex2);
        }
        return -1L;
    }
}
