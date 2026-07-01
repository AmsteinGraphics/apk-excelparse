package com.example.apkexcelparse.ui;

import android.content.Context;
import android.graphics.Canvas;
import android.graphics.Paint;
import android.util.AttributeSet;
import android.view.View;

/**
 * Compact row of colored dots, one per criterion. Dot color encodes the mark bucket:
 * -1 unmarked, 0..4 for values 0.0, 0.25, 0.5, 0.75, 1.0.
 */
public class DotProgressView extends View {

    private int[] values;
    private int highlightIndex = -1; // index of the dot to highlight, or -1 for none

    private final Paint fillPaint = new Paint(Paint.ANTI_ALIAS_FLAG);
    private final Paint strokePaint = new Paint(Paint.ANTI_ALIAS_FLAG);

    // 0.0 black, 0.25 red, 0.5 orange, 0.75 yellow, 1.0 green.
    private static final int[] BUCKET_COLORS = {
            0xFF000000,
            0xFFEF5350,
            0xFFFB8C00,
            0xFFFDD835,
            0xFF66BB6A,
    };

    public DotProgressView(Context context) {
        super(context);
        init();
    }

    public DotProgressView(Context context, AttributeSet attrs) {
        super(context, attrs);
        init();
    }

    public DotProgressView(Context context, AttributeSet attrs, int defStyleAttr) {
        super(context, attrs, defStyleAttr);
        init();
    }

    private void init() {
        strokePaint.setStyle(Paint.Style.STROKE);
        strokePaint.setColor(0xFF000000);
        strokePaint.setStrokeWidth(dpToPx(0.75f));
    }

    public void setValues(int[] values) {
        this.values = values;
        invalidate();
    }

    public void setValueAt(int index, int bucket) {
        if (values == null || index < 0 || index >= values.length) return;
        values[index] = bucket;
        invalidate();
    }

    /** Highlight the dot at this index with a black ring behind it; -1 clears the highlight. */
    public void setHighlightIndex(int index) {
        this.highlightIndex = index;
        invalidate();
    }

    private float dpToPx(float dp) {
        return dp * getResources().getDisplayMetrics().density;
    }

    @Override
    protected void onDraw(Canvas canvas) {
        if (values == null || values.length == 0) return;
        int n = values.length;
        int w = getWidth();
        int h = getHeight();
        if (w == 0 || h == 0) return;
        float slot = (float) w / n;
        float radius = Math.min(slot * 0.42f, h * 0.34f);
        if (radius < 1f) radius = 1f;
        // Highlight ring: larger black circle drawn behind the current dot.
        float highlightRadius = Math.min(Math.min(slot * 0.5f, h * 0.5f), radius * 1.7f);
        float cy = h / 2f;
        for (int i = 0; i < n; i++) {
            float cx = slot * (i + 0.5f);
            if (i == highlightIndex) {
                fillPaint.setColor(0xFF000000);
                canvas.drawCircle(cx, cy, highlightRadius, fillPaint);
            }
            int v = values[i];
            if (v < 0 || v >= BUCKET_COLORS.length) {
                fillPaint.setColor(0xFFFFFFFF);
                canvas.drawCircle(cx, cy, radius, fillPaint);
                canvas.drawCircle(cx, cy, radius, strokePaint);
            } else {
                fillPaint.setColor(BUCKET_COLORS[v]);
                canvas.drawCircle(cx, cy, radius, fillPaint);
            }
        }
    }
}
