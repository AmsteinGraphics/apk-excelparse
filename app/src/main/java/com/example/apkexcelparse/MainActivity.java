package com.example.apkexcelparse;

import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;

import org.json.JSONObject;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class MainActivity extends AppCompatActivity {
    private final ExecutorService executor = Executors.newSingleThreadExecutor();
    private final Handler mainHandler = new Handler(Looper.getMainLooper());

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        TextView versionText = findViewById(R.id.versionText);
        versionText.setText("Current version: 1.0.0");

        Button updateButton = findViewById(R.id.updateButton);
        updateButton.setOnClickListener(v -> checkForUpdates());
    }

    private void checkForUpdates() {
        executor.execute(() -> {
            try {
                URL url = new URL("https://api.github.com/repos/AmsteinGraphics/apk-excelparse/releases/latest");
                HttpURLConnection connection = (HttpURLConnection) url.openConnection();
                connection.setRequestMethod("GET");
                connection.setRequestProperty("Accept", "application/vnd.github+json");

                int responseCode = connection.getResponseCode();
                if (responseCode != HttpURLConnection.HTTP_OK) {
                    throw new RuntimeException("GitHub request failed with code " + responseCode);
                }

                BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream()));
                StringBuilder builder = new StringBuilder();
                String line;
                while ((line = reader.readLine()) != null) {
                    builder.append(line);
                }
                reader.close();

                JSONObject release = new JSONObject(builder.toString());
                String tagName = release.optString("tag_name", "unknown");
                String htmlUrl = release.optString("html_url", "https://github.com/AmsteinGraphics/apk-excelparse/releases");

                mainHandler.post(() -> {
                    Intent intent = new Intent(Intent.ACTION_VIEW, Uri.parse(htmlUrl));
                    startActivity(intent);
                    Toast.makeText(this, "Latest release: " + tagName, Toast.LENGTH_SHORT).show();
                });
            } catch (Exception error) {
                mainHandler.post(() -> Toast.makeText(this, "Unable to check updates: " + error.getMessage(), Toast.LENGTH_LONG).show());
            }
        });
    }
}
