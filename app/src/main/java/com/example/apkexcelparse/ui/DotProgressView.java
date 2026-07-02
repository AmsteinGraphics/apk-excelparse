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
    private float[] scales;          // per-dot radius multiplier (coefficient-based); null = all 1.0
    private String[] labels;         // per-dot centred text (e.g. coefficient); null = no labels
    private float fixedSlotPx = 0f;  // when > 0, dots use this slot width and left-pack instead of filling getWidth()
    private int highlightIndex = -1; // index of the dot to highlight, or -1 for none

    // Coefficient digit text size — a fixed px size, identical regardless of the dot's scale.
    private static final float DIGIT_TEXT_DP = 12f;

    // Base (coef-2, scale 1.0) dot radius. Fixed in dp — NOT derived from row height — so a taller
    // row only creates headroom for large-coefficient dots; it never enlarges the coef-2 dot.
    // 11.2dp preserves the historical size (old formula was h·0.34 at a 33dp row).
    private static final float BASE_RADIUS_DP = 11.2f;

    private final Paint fillPaint = new Paint(Paint.ANTI_ALIAS_FLAG);
    private final Paint strokePaint = new Paint(Paint.ANTI_ALIAS_FLAG);
    private final Paint textPaint = new Paint(Paint.ANTI_ALIAS_FLAG);

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
        textPaint.setColor(0xFFFFFFFF);
        textPaint.setTextAlign(Paint.Align.CENTER);
        textPaint.setFakeBoldText(true);
    }

    public void setValues(int[] values) {
        this.values = values;
        invalidate();
    }

    /** Per-dot radius multipliers, parallel to the values array. null = every dot at 1.0. */
    public void setScales(float[] scales) {
        this.scales = scales;
        invalidate();
    }

    /** Per-dot centred labels (parallel to values). Drawn only on graded dots. null = no labels. */
    public void setLabels(String[] labels) {
        this.labels = labels;
        invalidate();
    }

    /**
     * Force a fixed slot width (px) per dot so several DotProgressViews render dots at an
     * identical physical size regardless of their own width. 0 = fill the view width (default).
     */
    public void setFixedSlotPx(float slotPx) {
        this.fixedSlotPx = slotPx;
        requestLayout();
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

    // Highlight ring thickness (px the black circle extends beyond the dot), constant across scales.
    private static final float HIGHLIGHT_OVERSHOOT_DP = 3f;

    @Override
    protected void onMeasure(int widthMeasureSpec, int heightMeasureSpec) {
        if (fixedSlotPx > 0f && values != null && values.length > 0) {
            // Honour an exact width (fixed table cell) so dots left-pack in a uniform column;
            // otherwise size to the dot string so the view wraps its content.
            int wMode = MeasureSpec.getMode(widthMeasureSpec);
            int w = wMode == MeasureSpec.EXACTLY
                    ? MeasureSpec.getSize(widthMeasureSpec)
                    : Math.round(fixedSlotPx * values.length);
            int hMode = MeasureSpec.getMode(heightMeasureSpec);
            int h = hMode == MeasureSpec.UNSPECIFIED
                    ? Math.round(dpToPx(33f))
                    : MeasureSpec.getSize(heightMeasureSpec);
            setMeasuredDimension(w, h);
        } else {
            super.onMeasure(widthMeasureSpec, heightMeasureSpec);
        }
    }

    /** Draw a centred coefficient digit at the given colour, fixed px size regardless of dot scale. */
    private void drawLabel(Canvas canvas, float cx, float cy, String text, int color) {
        if (text == null) return;
        textPaint.setColor(color);
        textPaint.setTextSize(dpToPx(DIGIT_TEXT_DP));
        Paint.FontMetrics fm = textPaint.getFontMetrics();
        float baseline = cy - (fm.ascent + fm.descent) / 2f;
        canvas.drawText(text, cx, baseline, textPaint);
    }

    @Override
    protected void onDraw(Canvas canvas) {
        if (values == null || values.length == 0) return;
        int n = values.length;
        int w = getWidth();
        int h = getHeight();
        if (w == 0 || h == 0) return;
        float slot = fixedSlotPx > 0f ? fixedSlotPx : (float) w / n;
        float overshoot = dpToPx(HIGHLIGHT_OVERSHOOT_DP);
        // Leave room so the highlight ring (dot + fixed overshoot) never spills out of the row.
        float maxRadius = h * 0.5f - overshoot;
        // Base radius is a fixed dp (coef-2 size), still capped by the slot so crowded rows shrink.
        float baseRadius = Math.min(slot * 0.42f, dpToPx(BASE_RADIUS_DP));
        float cy = h / 2f;
        for (int i = 0; i < n; i++) {
            float cx = slot * (i + 0.5f);
            float scale = (scales != null && i < scales.length) ? scales[i] : 1f;
            float radius = baseRadius * scale;
            // Keep dots (and the ring) inside the row; horizontal overlap is intentional (weight cue).
            radius = Math.min(radius, maxRadius);
            if (radius < 1f) radius = 1f;
            if (i == highlightIndex) {
                // Highlight ring: a black circle a fixed number of px larger than the dot, so it
                // reads as a constant-weight stroke regardless of the dot's coefficient scale.
                fillPaint.setColor(0xFF000000);
                canvas.drawCircle(cx, cy, radius + overshoot, fillPaint);
            }
            int v = values[i];
            String label = (labels != null && i < labels.length) ? labels[i] : null;
            if (v < 0 || v >= BUCKET_COLORS.length) {
                fillPaint.setColor(0xFFFFFFFF);
                canvas.drawCircle(cx, cy, radius, fillPaint);
                canvas.drawCircle(cx, cy, radius, strokePaint);
                // Coefficient digit on blank dots too, in 50% gray.
                drawLabel(canvas, cx, cy, label, 0xFF808080);
            } else {
                fillPaint.setColor(BUCKET_COLORS[v]);
                canvas.drawCircle(cx, cy, radius, fillPaint);
                // Coefficient digit, white, fixed size — on graded dots.
                drawLabel(canvas, cx, cy, label, 0xFFFFFFFF);
            }
        }
    }
}
