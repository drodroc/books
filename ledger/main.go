//go:build windows

package main

import (
	"embed"
	"io"
	"net"
	"net/http"
	"os"
	"path/filepath"
	"strconv"

	"github.com/jchv/go-webview2"
	"golang.org/x/sys/windows"
)

//go:embed index.html app.js core.js styles.css lib
var assets embed.FS

func main() {
	dataPath, err := localDataDir()
	if err != nil {
		showError("数据文件夹没有建在程序旁边：" + err.Error())
		return
	}
	port := choosePort(dataPath)
	listener, err := net.Listen("tcp", "127.0.0.1:"+strconvPort(port))
	if err != nil {
		showError("月度算账已经在运行，或原来的账本端口被占用。请先关掉已经打开的窗口，再重新打开。")
		return
	}
	go http.Serve(listener, http.FileServer(http.FS(assets)))

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

func localDataDir() (string, error) {
	exe, err := os.Executable()
	if err != nil {
		return "", err
	}
	if resolved, err := filepath.EvalSymlinks(exe); err == nil {
		exe = resolved
	}
	// 「数据」这个目录名不能改。换掉 exe 后还是读这里，已经配对的账会留着。
	dir := filepath.Join(filepath.Dir(exe), "数据")
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return "", err
	}
	legacy := legacyDataDir()
	if !profileExists(dir) {
		_ = copyTree(legacy, dir)
	} else if findBestOrigin(legacy).size > findBestOrigin(dir).size {
		_ = copyTree(legacy, dir)
	}
	return dir, nil
}

func choosePort(dataDir string) int {
	hit := findBestOrigin(dataDir, legacyDataDir())
	if hit.port > 0 && hit.size > 0 {
		return hit.port
	}
	return stablePort
}

func strconvPort(port int) string {
	return strconv.Itoa(port)
}

func legacyDataDir() string {
	base := os.Getenv("APPDATA")
	if base == "" {
		if dir, err := os.UserConfigDir(); err == nil {
			base = dir
		}
	}
	if base == "" {
		return ""
	}
	return filepath.Join(base, "月度算账")
}

func profileExists(dir string) bool {
	info, err := os.Stat(filepath.Join(dir, "EBWebView"))
	return err == nil && info.IsDir()
}

func copyTree(src, dst string) error {
	if src == "" {
		return nil
	}
	info, err := os.Stat(src)
	if err != nil || !info.IsDir() {
		return nil
	}
	return filepath.Walk(src, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		rel, err := filepath.Rel(src, path)
		if err != nil {
			return err
		}
		target := filepath.Join(dst, rel)
		if info.IsDir() {
			return os.MkdirAll(target, 0o755)
		}
		if err := os.MkdirAll(filepath.Dir(target), 0o755); err != nil {
			return err
		}
		in, err := os.Open(path)
		if err != nil {
			return err
		}
		defer in.Close()
		out, err := os.Create(target)
		if err != nil {
			return err
		}
		_, copyErr := io.Copy(out, in)
		closeErr := out.Close()
		if copyErr != nil {
			return copyErr
		}
		return closeErr
	})
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
