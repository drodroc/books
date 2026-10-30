//go:build windows

package main

import (
	"embed"
	"net"
	"net/http"
	"os"
	"path/filepath"

	"github.com/jchv/go-webview2"
	"golang.org/x/sys/windows"
)

//go:embed index.html app.js core.js styles.css lib
var assets embed.FS

func main() {
	listener, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		showError("算账窗口没有打开：" + err.Error())
		return
	}
	go http.Serve(listener, http.FileServer(http.FS(assets)))

	dataPath := filepath.Join(os.Getenv("APPDATA"), "月度算账")
	if os.Getenv("APPDATA") == "" {
		dir, dirErr := os.UserConfigDir()
		if dirErr == nil {
			dataPath = filepath.Join(dir, "月度算账")
		}
	}
	_ = os.MkdirAll(dataPath, 0o755)

	view := webview2.NewWithOptions(webview2.WebViewOptions{
		Debug:     false,
		DataPath:  dataPath,
		AutoFocus: true,
		WindowOptions: webview2.WindowOptions{
			Title:  "月度算账",
			Width:  1440,
			Height: 900,
			Center: true,
		},
	})
	if view == nil {
		showError("这台电脑还没有打开窗口所需的 Microsoft Edge WebView2。到微软官网安装 WebView2 运行库后，再双击本程序。")
		return
	}
	defer view.Destroy()
	view.Navigate("http://" + listener.Addr().String() + "/index.html")
	view.Run()
}

func showError(message string) {
	text, err := windows.UTF16PtrFromString(message)
	if err != nil {
		return
	}
	title, err := windows.UTF16PtrFromString("月度算账")
	if err != nil {
		return
	}
	windows.MessageBox(0, text, title, windows.MB_ICONERROR|windows.MB_OK)
}
