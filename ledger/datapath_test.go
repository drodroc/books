package main

import (
	"os"
	"path/filepath"
	"testing"
)

func TestFindBestOriginPrefersLargerDatabase(t *testing.T) {
	root := t.TempDir()
	small := filepath.Join(root, "EBWebView", "Default", "IndexedDB", "http_127.0.0.1_11111.indexeddb.leveldb")
	large := filepath.Join(root, "EBWebView", "Default", "IndexedDB", "http_127.0.0.1_22222.indexeddb.leveldb")
	if err := os.MkdirAll(small, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(large, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(small, "a"), []byte("x"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(large, "a"), []byte("this database has the real books"), 0o644); err != nil {
		t.Fatal(err)
	}
	hit := findBestOrigin(root)
	if hit.port != 22222 {
		t.Fatalf("port = %d, want 22222", hit.port)
	}
}

func TestFindBestOriginAcrossOldAndNewFolders(t *testing.T) {
	oldRoot := t.TempDir()
	newRoot := t.TempDir()
	oldDB := filepath.Join(oldRoot, "EBWebView", "Default", "IndexedDB", "http_127.0.0.1_34567.indexeddb.leveldb")
	if err := os.MkdirAll(oldDB, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(oldDB, "data"), []byte("configured income and expenses"), 0o644); err != nil {
		t.Fatal(err)
	}
	hit := findBestOrigin(newRoot, oldRoot)
	if hit.port != 34567 || hit.size == 0 {
		t.Fatalf("hit = %+v", hit)
	}
}
