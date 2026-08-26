package com.carmaintenance.app;

import android.annotation.SuppressLint;
import android.app.AlertDialog;
import android.content.SharedPreferences;
import android.graphics.Bitmap;
import android.os.Bundle;
import android.view.KeyEvent;
import android.view.View;
import android.webkit.WebChromeClient;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.EditText;
import android.widget.Toast;
import androidx.appcompat.app.AppCompatActivity;
import androidx.swiperefreshlayout.widget.SwipeRefreshLayout;

public class MainActivity extends AppCompatActivity {

    private WebView webView;
    private SwipeRefreshLayout swipeRefresh;
    private SharedPreferences prefs;
    private static final String PREF_SERVER_URL = "server_url";
    private static final String DEFAULT_URL = "http://192.168.1.100:9595";

    @SuppressLint("SetJavaScriptEnabled")
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        prefs = getSharedPreferences("app_settings", MODE_PRIVATE);
        String serverUrl = prefs.getString(PREF_SERVER_URL, DEFAULT_URL);

        webView = findViewById(R.id.webView);
        swipeRefresh = findViewById(R.id.swipeRefresh);

        WebSettings settings = webView.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        settings.setDatabaseEnabled(true);
        settings.setLoadWithOverviewMode(true);
        settings.setUseWideViewPort(true);
        settings.setBuiltInZoomControls(false);
        settings.setDisplayZoomControls(false);

        webView.setWebViewClient(new WebViewClient() {
            @Override
            public void onPageStarted(WebView view, String url, Bitmap favicon) {
                swipeRefresh.setRefreshing(true);
            }

            @Override
            public void onPageFinished(WebView view, String url) {
                swipeRefresh.setRefreshing(false);
            }

            @Override
            public void onReceivedError(WebView view, int errorCode, String description, String failingUrl) {
                swipeRefresh.setRefreshing(false);
                showConnectionErrorDialog();
            }
        });

        webView.setWebChromeClient(new WebChromeClient());

        swipeRefresh.setOnRefreshListener(() -> webView.reload());

        // Long click to configure server URL
        webView.setOnLongClickListener(v -> {
            showServerConfigDialog();
            return false;
        });

        webView.loadUrl(serverUrl);
    }

    private void showServerConfigDialog() {
        AlertDialog.Builder builder = new AlertDialog.Builder(this);
        builder.setTitle("Настройка адреса сервера");
        builder.setMessage("Введите IP-адрес и порт вашего сервера (например: http://192.168.1.150:9595)");

        final EditText input = new EditText(this);
        input.setText(prefs.getString(PREF_SERVER_URL, DEFAULT_URL));
        builder.setView(input);

        builder.setPositiveButton("Сохранить", (dialog, which) -> {
            String newUrl = input.getText().toString().trim();
            if (!newUrl.startsWith("http://") && !newUrl.startsWith("https://")) {
                newUrl = "http://" + newUrl;
            }
            prefs.edit().putString(PREF_SERVER_URL, newUrl).apply();
            webView.loadUrl(newUrl);
            Toast.makeText(this, "Адрес обновлен: " + newUrl, Toast.LENGTH_SHORT).show();
        });

        builder.setNegativeButton("Отмена", (dialog, which) -> dialog.cancel());
        builder.show();
    }

    private void showConnectionErrorDialog() {
        new AlertDialog.Builder(this)
            .setTitle("Ошибка подключения")
            .setMessage("Не удалось подключиться к серверу: " + prefs.getString(PREF_SERVER_URL, DEFAULT_URL) + "\n\nПроверьте, что сервер запущен, или укажите правильный IP-адрес сервера.")
            .setPositiveButton("Сменить адрес", (dialog, which) -> showServerConfigDialog())
            .setNegativeButton("Повторить", (dialog, which) -> webView.reload())
            .show();
    }

    @Override
    public boolean onKeyDown(int keyCode, KeyEvent event) {
        if (keyCode == KeyEvent.KEYCODE_BACK && webView.canGoBack()) {
            webView.goBack();
            return true;
        }
        return super.onKeyDown(keyCode, event);
    }
}
