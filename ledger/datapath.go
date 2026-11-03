package main

import (
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

const stablePort = 18765

var originPort = regexp.MustCompile(`(?i)(?:127\.0\.0\.1|localhost)[_-](\d{2,5})`)

type originHit struct {
	port int
	size int64
}

func findBestOrigin(roots ...string) originHit {
	var best originHit
	for _, root := range roots {
		if root == "" {
			continue
		}
		_ = filepath.Walk(root, func(path string, info os.FileInfo, err error) error {
			if err != nil || info == nil || !info.IsDir() {
				return nil
			}
			name := info.Name()
			if !strings.Contains(strings.ToLower(name), "indexeddb") {
				return nil
			}
			match := originPort.FindStringSubmatch(name)
			if match == nil {
				return nil
			}
			port, err := strconv.Atoi(match[1])
			if err != nil || port < 1 || port > 65535 {
				return nil
			}
			size := dirSize(path)
			if size > best.size {
				best = originHit{port: port, size: size}
			}
			return nil
		})
	}
	return best
}

func dirSize(root string) int64 {
	var total int64
	_ = filepath.Walk(root, func(_ string, info os.FileInfo, err error) error {
		if err == nil && info != nil && !info.IsDir() {
			total += info.Size()
		}
		return nil
	})
	return total
}
